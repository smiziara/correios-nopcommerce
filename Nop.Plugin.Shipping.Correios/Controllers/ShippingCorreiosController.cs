using System;
using System.Text;
using System.Web.Mvc;
using Nop.Plugin.Shipping.Correios.Domain;
using Nop.Plugin.Shipping.Correios.Models;
using Nop.Services.Configuration;
using Nop.Web.Framework.Controllers;
using Nop.Core;
using Nop.Services.Localization;

namespace Nop.Plugin.Shipping.Correios.Controllers
{
	[AdminAuthorize]
	public class ShippingCorreiosController : Controller
	{
		private readonly CorreiosSettings _correiosSettings;
		private readonly ISettingService _settingService;
		private readonly ILocalizationService _localizationService;

		public ShippingCorreiosController(CorreiosSettings correiosSettings, ISettingService settingService, ILocalizationService localizationService)
		{
			this._correiosSettings = correiosSettings;
			this._settingService = settingService;
			this._localizationService = localizationService;
		}

		public ActionResult Configure()
		{
			var model = new CorreiosShippingModel();
			model.Url = _correiosSettings.Url;
			model.CodigoEmpresa = _correiosSettings.CodigoEmpresa;
			model.Senha = _correiosSettings.Senha;
			model.CustoAdicionalEnvio = _correiosSettings.CustoAdicionalEnvio;
			model.IncluirAvisoRecebimento = _correiosSettings.IncluirAvisoRecebimento;
			model.IncluirMaoPropria = _correiosSettings.IncluirMaoPropria;
			model.IncluirValorDeclarado = _correiosSettings.IncluirValorDeclarado;
			model.DiasUteisAdicionais = _correiosSettings.DiasUteisAdicionais;

			var services = new CorreiosServices();
			// Load service names
			string carrierServicesOfferedDomestic = _correiosSettings.CarrierServicesOffered;
			foreach (string service in services.Services)
				model.AvailableCarrierServices.Add(service);

			if (!String.IsNullOrEmpty(carrierServicesOfferedDomestic))
				foreach (string service in services.Services)
				{
					string serviceId = CorreiosServices.GetServiceId(service);
					if (!String.IsNullOrEmpty(serviceId) && !String.IsNullOrEmpty(carrierServicesOfferedDomestic))
					{
						if (carrierServicesOfferedDomestic.Contains(serviceId))
							model.CarrierServicesOffered.Add(service);
					}
				}


			return View("Nop.Plugin.Shipping.Correios.Views.ShippingCorreios.Configure", model);
		}

		[HttpPost]
		public ActionResult Configure(CorreiosShippingModel model)
		{
			if (!ModelState.IsValid)
			{
				return Configure();
			}

			//save settings
			 _correiosSettings.Url = model.Url ;
			 _correiosSettings.CodigoEmpresa = model.CodigoEmpresa;
			 _correiosSettings.Senha = model.Senha;
			 _correiosSettings.CustoAdicionalEnvio = model.CustoAdicionalEnvio;
			 _correiosSettings.IncluirAvisoRecebimento = model.IncluirAvisoRecebimento;
			 _correiosSettings.IncluirMaoPropria = model.IncluirMaoPropria;
			 _correiosSettings.IncluirValorDeclarado = model.IncluirValorDeclarado;
			 _correiosSettings.DiasUteisAdicionais = model.DiasUteisAdicionais;
			// Save selected services
			var carrierServicesOfferedDomestic = new StringBuilder();
			int carrierServicesDomesticSelectedCount = 0;
			if (model.CheckedCarrierServices != null)
			{
				foreach (var cs in model.CheckedCarrierServices)
				{
					carrierServicesDomesticSelectedCount++;
					string serviceId = CorreiosServices.GetServiceId(cs);
					if (!String.IsNullOrEmpty(serviceId))
						carrierServicesOfferedDomestic.AppendFormat("{0},", serviceId);
				}
			}

			if (carrierServicesDomesticSelectedCount == 0)
				_correiosSettings.CarrierServicesOffered = "41106,40010,40215";
			else
				_correiosSettings.CarrierServicesOffered = carrierServicesOfferedDomestic.ToString().TrimEnd(',');


			_settingService.SaveSetting(_correiosSettings);

			ViewData["sucesso"] = this._localizationService.GetResource("Admin.Configuration.Updated");

			return Configure();
		}
	}
}
