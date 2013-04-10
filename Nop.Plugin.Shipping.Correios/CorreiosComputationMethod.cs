using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Web.Routing;
using System.Web.Services.Protocols;
using Nop.Core;
using Nop.Core.Domain.Directory;
using Nop.Core.Domain.Discounts;
using Nop.Core.Domain.Shipping;
using Nop.Core.Plugins;
using Nop.Plugin.Shipping.Correios.Domain;
using Nop.Plugin.Shipping.Correios.CalcPrecoPrazoWebReference;
using Nop.Services.Configuration;
using Nop.Services.Directory;
using Nop.Services.Orders;
using Nop.Services.Shipping;
using Nop.Services.Logging;
using System.Linq;
using System.Globalization;
using Nop.Services.Common;

namespace Nop.Plugin.Shipping.Correios
{
	/// <summary>
	/// Correios computation method.
	/// </summary>
	public class CorreiosComputationMethod : BasePlugin, IShippingRateComputationMethod
	{
		#region Constants
		private const int MAX_PACKAGE_WEIGHT = 30;
		private const int MAX_PACKAGE_TOTAL_DIMENSION = 200;
		private const int MAX_PACKAGE_SIZES = 105;

		private const int MIN_PACKAGE_LENGTH = 16;
		private const int MIN_PACKAGE_WIDTH = 11;
		private const int MIN_PACKAGE_HEIGHT = 2;
		private const int MIN_PACKAGE_SIZE = 29;

		private const int MAX_ROLL_TOTAL_DIMENSION = 200;
		private const int MAX_ROLL_LENGTH = 105;
		private const int MAX_ROLL_DIAMETER = 91;

		private const int MIN_ROLL_LENGTH = 18;
		private const int MIN_ROLL_DIAMETER = 5;
		private const int MIN_ROLL_SIZE = 28;

		private const string MEASURE_WEIGHT_SYSTEM_KEYWORD = "kg";
		private const string MEASURE_DIMENSION_SYSTEM_KEYWORD = "millimetres";
		#endregion

		#region Fields
		private readonly IMeasureService _measureService;
		private readonly IShippingService _shippingService;
		private readonly ISettingService _settingService;
		private readonly CorreiosSettings _correiosSettings;
		private readonly IOrderTotalCalculationService _orderTotalCalculationService;
		private readonly ICurrencyService _currencyService;
		private readonly CurrencySettings _currencySettings;
		private readonly ShippingSettings _shippingSettings;
		private readonly IAddressService _addressService;
		private readonly ILogger _logger;
		#endregion

		#region Ctor
		public CorreiosComputationMethod(IMeasureService measureService,
			IShippingService shippingService, ISettingService settingService,
			CorreiosSettings correiosSettings, IOrderTotalCalculationService orderTotalCalculationService,
			ICurrencyService currencyService, CurrencySettings currencySettings, ShippingSettings shippingSettings, IAddressService addressService, ILogger logger)
		{
			this._measureService = measureService;
			this._shippingService = shippingService;
			this._settingService = settingService;
			this._correiosSettings = correiosSettings;
			this._orderTotalCalculationService = orderTotalCalculationService;
			this._currencyService = currencyService;
			this._currencySettings = currencySettings;
			this._shippingSettings = shippingSettings;
			this._addressService = addressService;
			this._logger = logger;
		}
		#endregion

		#region Utilities
		private cResultado ProcessShipping(GetShippingOptionRequest getShippingOptionRequest, GetShippingOptionResponse getShippingOptionResponse)
		{
			var usedMeasureWeight = _measureService.GetMeasureWeightBySystemKeyword(MEASURE_WEIGHT_SYSTEM_KEYWORD);

			if (usedMeasureWeight == null)
			{
				string e = string.Format("Plugin.Shipping.Correios: Could not load \"{0}\" measure weight", MEASURE_WEIGHT_SYSTEM_KEYWORD);

				_logger.Fatal(e);

				throw new NopException(e);
			}

			var usedMeasureDimension = _measureService.GetMeasureDimensionBySystemKeyword(MEASURE_DIMENSION_SYSTEM_KEYWORD);

			if (usedMeasureDimension == null)
			{
				string e = string.Format("Plugin.Shipping.Correios: Could not load \"{0}\" measure dimension", MEASURE_DIMENSION_SYSTEM_KEYWORD);

				_logger.Fatal(e);

				throw new NopException(e);
			}


			//Na versão 2.2 o getShippingOptionRequest.ZipPostalCodeFrom retorna string.Empty, possui um TODO...

			string cepOrigem = null;

			if (this._shippingSettings.ShippingOriginAddressId > 0)
			{
				var addr = this._addressService.GetAddressById(this._shippingSettings.ShippingOriginAddressId);

				if (addr != null && !String.IsNullOrEmpty(addr.ZipPostalCode) && addr.ZipPostalCode.Length >= 8 && addr.ZipPostalCode.Length <= 9)
				{
					cepOrigem = addr.ZipPostalCode;
				}
			}

			if (cepOrigem == null)
			{
				_logger.Fatal("Plugin.Shipping.Correios: CEP de Envio em branco ou inválido, configure nas opções de envio do NopCommerce.Em Administração > Configurações > Configurações de Envio. Formato: 00000000");

				throw new NopException("Plugin.Shipping.Correios: CEP de Envio em branco ou inválido, configure nas opções de envio do NopCommerce.Em Administração > Configurações > Configurações de Envio. Formato: 00000000");
			}
			
			var correiosCalculation = new CorreiosBatchCalculation(this._logger)
			{
				CodigoEmpresa = _correiosSettings.CodigoEmpresa,
				Senha = _correiosSettings.Senha,
				CepOrigem = cepOrigem,
				Servicos = _correiosSettings.CarrierServicesOffered,
				AvisoRecebimento = _correiosSettings.IncluirAvisoRecebimento,
				MaoPropria = _correiosSettings.IncluirMaoPropria,
				CepDestino = getShippingOptionRequest.ShippingAddress.ZipPostalCode
			};

			decimal subtotalBase = decimal.Zero;
			decimal orderSubTotalDiscountAmount = decimal.Zero;
			Discount orderSubTotalAppliedDiscount = null;
			decimal subTotalWithoutDiscountBase = decimal.Zero;
			decimal subTotalWithDiscountBase = decimal.Zero;

			_orderTotalCalculationService.GetShoppingCartSubTotal(getShippingOptionRequest.Items,
				out orderSubTotalDiscountAmount, out orderSubTotalAppliedDiscount,
				out subTotalWithoutDiscountBase, out subTotalWithDiscountBase);

			subtotalBase = subTotalWithDiscountBase;

			int length = Convert.ToInt32(Math.Ceiling(_measureService.ConvertFromPrimaryMeasureDimension(getShippingOptionRequest.GetTotalLength(), usedMeasureDimension)) / 10);
			int height = Convert.ToInt32(Math.Ceiling(_measureService.ConvertFromPrimaryMeasureDimension(getShippingOptionRequest.GetTotalHeight(), usedMeasureDimension)) / 10);
			int width = Convert.ToInt32(Math.Ceiling(_measureService.ConvertFromPrimaryMeasureDimension(getShippingOptionRequest.GetTotalWidth(), usedMeasureDimension)) / 10);
			int weight = Convert.ToInt32(Math.Ceiling(_measureService.ConvertFromPrimaryMeasureWeight(_shippingService.GetShoppingCartTotalWeight(getShippingOptionRequest.Items), usedMeasureWeight)));

			if (length < 1)
				length = 1;
			if (height < 1)
				height = 1;
			if (width < 1)
				width = 1;
			if (weight < 1)
				weight = 1;

			//Altura não pode ser maior que o comprimento, para evitar erro, igualamos e a embalagem deve ser adaptada.
			if (height > length)
			{
				length = height;
			}

			if (IsPackageTooSmall(length, height, width))
			{
				length = MIN_PACKAGE_LENGTH;
				height = MIN_PACKAGE_HEIGHT;
				width = MIN_PACKAGE_WIDTH;
			}

			if ((!IsPackageTooHeavy(weight)) && (!IsPackageTooLarge(length, height, width)))
			{
				Debug.WriteLine("Plugin.Shipping.Correios: Pacote unico");

				correiosCalculation.Pacotes.Add(new CorreiosBatchCalculation.Pacote()
				{
					Altura = height,
					Comprimento = length,
					Largura = width,
					Diametro = 0,
					FormatoPacote = true,
					Peso = weight,
					ValorDeclarado = (_correiosSettings.IncluirValorDeclarado ? subtotalBase : 0)
				});

				return correiosCalculation.Calculate();
			}
			else
			{
				int totalPackages = 1;
				int totalPackagesDims = 1;
				int totalPackagesWeights = 1;

				if (IsPackageTooHeavy(weight))
				{
					totalPackagesWeights = Convert.ToInt32(Math.Ceiling((decimal)weight / (decimal)MAX_PACKAGE_WEIGHT));
				}
				if (IsPackageTooLarge(length, height, width))
				{
					totalPackagesDims = Convert.ToInt32(Math.Ceiling((decimal)TotalPackageSize(length, height, width) / (decimal)MAX_PACKAGE_TOTAL_DIMENSION));
				}
				totalPackages = totalPackagesDims > totalPackagesWeights ? totalPackagesDims : totalPackagesWeights;

				if (totalPackages == 0)
					totalPackages = 1;

				int weight2 = weight / totalPackages;
				int height2 = height / totalPackages;
				int width2 = width / totalPackages;
				int length2 = length / totalPackages;

				if (weight2 < 1)
					weight2 = 1;
				if (height2 < 1)
					height2 = 1;
				if (width2 < 1)
					width2 = 1;
				if (length2 < 1)
					length2 = 1;

				//Altura não pode ser maior que o comprimento, para evitar erro, igualamos e a embalagem deve ser adaptada.
				if (height2 > width2)
					width2 = height2;

				Debug.WriteLine("Plugin.Shipping.Correios: Multiplos pacotes");

				correiosCalculation.Pacotes.Add(new CorreiosBatchCalculation.Pacote()
				{
					Altura = height2,
					Comprimento = length2,
					Largura = width2,
					Diametro = 0,
					FormatoPacote = true,
					Peso = weight2,
					ValorDeclarado = (_correiosSettings.IncluirValorDeclarado ? subtotalBase / totalPackages : 0)
				});

				var result = correiosCalculation.Calculate();

				if (result != null)
				{
					foreach (cServico s in result.Servicos)
					{
						if (s.Erro == "0")
						{
							s.Valor = (decimal.Parse(s.Valor, correiosCalculation.PtBrCulture) * totalPackages).ToString(correiosCalculation.PtBrCulture);
							s.ValorAvisoRecebimento = (decimal.Parse(s.ValorAvisoRecebimento, correiosCalculation.PtBrCulture) * totalPackages).ToString(correiosCalculation.PtBrCulture);
							s.ValorMaoPropria = (decimal.Parse(s.ValorMaoPropria, correiosCalculation.PtBrCulture) * totalPackages).ToString(correiosCalculation.PtBrCulture);
							s.ValorValorDeclarado = (decimal.Parse(s.ValorValorDeclarado, correiosCalculation.PtBrCulture) * totalPackages).ToString(correiosCalculation.PtBrCulture);
						}
					}
				}

				return result;
			}
		}

		private bool IsPackageTooLarge(int length, int height, int width)
		{
			int total = TotalPackageSize(length, height, width);

			if (total > MAX_PACKAGE_TOTAL_DIMENSION || length > MAX_PACKAGE_SIZES || height > MAX_PACKAGE_SIZES || width > MAX_PACKAGE_SIZES)
				return true;
			else
				return false;
		}

		private bool IsPackageTooSmall(int length, int height, int width)
		{
			int total = TotalPackageSize(length, height, width);

			if (total < MIN_PACKAGE_SIZE || length < MIN_PACKAGE_LENGTH || height < MIN_PACKAGE_HEIGHT || width < MIN_PACKAGE_WIDTH)
				return true;
			else
				return false;
		}

		private int TotalPackageSize(int length, int height, int width)
		{
			return length + width + height;
		}

		private bool IsPackageTooHeavy(int weight)
		{
			if (weight > MAX_PACKAGE_WEIGHT)
				return true;
			else
				return false;
		}

		private bool IsRollTooLarge(int length, int diameter)
		{
			int total = TotalRollSize(length, diameter);

			if (total > MAX_ROLL_TOTAL_DIMENSION || length > MAX_ROLL_LENGTH || diameter > MAX_ROLL_DIAMETER)
				return true;
			else
				return false;
		}

		private bool IsRollTooSmall(int length, int diameter)
		{
			int total = TotalRollSize(length, diameter);

			if (total < MIN_ROLL_SIZE || length < MIN_ROLL_LENGTH || diameter < MIN_ROLL_DIAMETER)
				return true;
			else
				return false;
		}

		private int TotalRollSize(int length, int diameter)
		{
			return length + 2 * diameter;
		}

		private bool IsRollTooHeavy(int weight)
		{
			if (weight > MAX_PACKAGE_WEIGHT)
				return true;
			else
				return false;
		}
		#endregion

		#region Methods
		/// <summary>
		///  Gets available shipping options
		/// </summary>
		/// <param name="getShippingOptionRequest">A request for getting shipping options</param>
		/// <returns>Represents a response of getting shipping rate options</returns>
		public GetShippingOptionResponse GetShippingOptions(GetShippingOptionRequest getShippingOptionRequest)
		{
			if (getShippingOptionRequest == null)
				throw new ArgumentNullException("getShippingOptionRequest");

			var response = new GetShippingOptionResponse();

			if (getShippingOptionRequest.Items == null)
			{
				response.AddError("Sem items para enviar");
				return response;
			}

			if (getShippingOptionRequest.ShippingAddress == null)
			{
				response.AddError("Endereço de envio em branco");
				return response;
			}

			if (getShippingOptionRequest.ShippingAddress.ZipPostalCode == null)
			{
				response.AddError("CEP de envio em branco");
				return response;
			}

			var result = ProcessShipping(getShippingOptionRequest, response);

			if (result == null)
			{
				response.AddError("Não há serviços disponíveis no momento");
				return response;
			}
			else
			{
				List<string> group = new List<string>();

				foreach (cServico servico in result.Servicos.OrderBy(s => s.Valor))
				{
					Debug.WriteLine("Plugin.Shipping.Correios: Retorno WS");
					Debug.WriteLine("Codigo: " + servico.Codigo);
					Debug.WriteLine("Valor: " + servico.Valor);
					Debug.WriteLine("Valor Mão Própria: " + servico.ValorMaoPropria);
					Debug.WriteLine("Valor Aviso Recebimento: " + servico.ValorAvisoRecebimento);
					Debug.WriteLine("Valor Declarado: " + servico.ValorValorDeclarado);
					Debug.WriteLine("Prazo Entrega: " + servico.PrazoEntrega);
					Debug.WriteLine("Entrega Domiciliar: " + servico.EntregaDomiciliar);
					Debug.WriteLine("Entrega Sabado: " + servico.EntregaSabado);
					Debug.WriteLine("Erro: " + servico.Erro);
					Debug.WriteLine("Msg Erro: " + servico.MsgErro);

					if (servico.Erro == "0")
					{
						string name = CorreiosServices.GetServicePublicNameById(servico.Codigo.ToString());

						if (!group.Contains(name))
						{
							ShippingOption option = new ShippingOption();
							option.Name = name;
							option.Description = "Prazo médio de entrega " + (servico.PrazoEntrega + _correiosSettings.DiasUteisAdicionais) + " dias úteis";
							option.Rate = decimal.Parse(servico.Valor, CultureInfo.GetCultureInfo("pt-BR")) + _orderTotalCalculationService.GetShoppingCartAdditionalShippingCharge(getShippingOptionRequest.Items) + _correiosSettings.CustoAdicionalEnvio;
							response.ShippingOptions.Add(option);

							group.Add(name);
						}
					}
					else
					{
						_logger.Error("Plugin.Shipping.Correios: erro ao calcular frete: (" + CorreiosServices.GetServiceName(servico.Codigo.ToString()) + ")( " + servico.Erro + ") " + servico.MsgErro);
					}
				}

				return response;
			}
		}

		/// <summary>
		/// Gets fixed shipping rate (if shipping rate computation method allows it and the rate can be calculated before checkout).
		/// </summary>
		/// <param name="getShippingOptionRequest">A request for getting shipping options</param>
		/// <returns>Fixed shipping rate; or null in case there's no fixed shipping rate</returns>
		public decimal? GetFixedRate(GetShippingOptionRequest getShippingOptionRequest)
		{
			return null;
		}

		/// <summary>
		/// Gets a route for provider configuration
		/// </summary>
		/// <param name="actionName">Action name</param>
		/// <param name="controllerName">Controller name</param>
		/// <param name="routeValues">Route values</param>
		public void GetConfigurationRoute(out string actionName, out string controllerName, out RouteValueDictionary routeValues)
		{
			actionName = "Configure";
			controllerName = "ShippingCorreios";
			routeValues = new RouteValueDictionary() { { "Namespaces", "Nop.Plugin.Shipping.Correios.Controllers" }, { "area", null } };
		}

		/// <summary>
		/// Install plugin
		/// </summary>
		public override void Install()
		{
			var settings = new CorreiosSettings()
			{
				Url = "http://ws.correios.com.br/calculador/CalcPrecoPrazo.asmx",
				CodigoEmpresa = String.Empty,
				Senha = String.Empty
			};

			_settingService.SaveSetting(settings);

			base.Install();
		}
		#endregion

		#region Properties
		/// <summary>
		/// Gets a shipping rate computation method type
		/// </summary>
		public ShippingRateComputationMethodType ShippingRateComputationMethodType
		{
			get
			{
				return ShippingRateComputationMethodType.Realtime;
			}
		}
		#endregion

		public Services.Shipping.Tracking.IShipmentTracker ShipmentTracker
		{
			get { return new CorreiosShipmentTracker(this._logger, this._correiosSettings); }
		}
	}
}
