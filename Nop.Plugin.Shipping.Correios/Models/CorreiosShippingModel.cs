using System.Collections.Generic;
using System.ComponentModel;

namespace Nop.Plugin.Shipping.Correios.Models
{
	public class CorreiosShippingModel
	{
		public CorreiosShippingModel()
		{
			CarrierServicesOffered = new List<string>();
			AvailableCarrierServices = new List<string>();
		}

		[DisplayName("URL")]
		public string Url { get; set; }

		[DisplayName("Código da Empresa")]
		public string CodigoEmpresa { get; set; }

		[DisplayName("Senha")]
		public string Senha { get; set; }

		[DisplayName("Custo adicional de envio")]
		public decimal CustoAdicionalEnvio { get; set; }

		[DisplayName("Mão Própria?")]
		public bool IncluirMaoPropria { get; set; }

		[DisplayName("Valor Declarado?")]
		public bool IncluirValorDeclarado { get; set; }

		[DisplayName("Aviso de Recebimento?")]
		public bool IncluirAvisoRecebimento { get; set; }

		[DisplayName("Dias uteis adicionais ao prazo")]
		public int DiasUteisAdicionais { get; set; }

		public IList<string> CarrierServicesOffered { get; set; }
		public IList<string> AvailableCarrierServices { get; set; }
		public string[] CheckedCarrierServices { get; set; }
	}
}
