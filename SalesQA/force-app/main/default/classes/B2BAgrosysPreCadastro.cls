public without sharing class B2BAgrosysPreCadastro extends B2BAgrosys {

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpCliente> tp_cliente;
    }

    public String setCustomer(Lead customer) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'SET-PRECADCLI';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_cliente = new List<TpCliente>();

        TpCliente TpCliente = new TpCliente();
        TpCliente.ind_underline = true;
        TpCliente.razsoc = customer.Company;
        TpCliente.nome_fantasia = customer.NomeFantasia__c;
        TpCliente.cnpj = customer.CNPJInput__c;
        TpCliente.fone = customer.Phone;
        TpCliente.ram_codigo = Integer.valueOf(customer.RamoAtividade__c);
        TpCliente.cep = customer.PostalCode;
        TpCliente.endereco = customer.Street;
        TpCliente.numero = customer.NumeroEndereco__c;
        TpCliente.bairro = customer.BairroEndereco__c;
        TpCliente.nom_cidade = customer.City;
        TpCliente.estado = customer.State;
        TpCliente.complem = customer.ComplementoEndereco__c;

        dadosReceb.tp_cliente.add(TpCliente);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'setCustomer';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysPreCadastro - ' + customer.CNPJInput__c;

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        
        if(retorno.tp_retorno != null && !retorno.tp_retorno.isEmpty() && retorno.tp_retorno[0].msg_codigo == 1) {
        	return 'success';
        }
        else {
        	return 'integração error';
        }
    }
}