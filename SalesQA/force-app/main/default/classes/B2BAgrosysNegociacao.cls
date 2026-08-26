public without sharing class B2BAgrosysNegociacao extends B2BAgrosys {
    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer cli_codigo;
    }
    
    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    @AuraEnabled(cacheable=false)
    public static Boolean getCustomerCode() {
        User usr = [
            SELECT Contact.AccountId, Contact.Account.CodigoClienteERP__c
            FROM User WHERE Id = :UserInfo.getUserId()
        ];
        
        Boolean hasNegociacao = usr.Contact.Account.CodigoClienteERP__c != null ? B2BAgrosysNegociacao.getNegociacao(usr) : null;
        return hasNegociacao;
    }

    static Boolean getNegociacao(User user) {
        Request dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-CLIENTE-NEGOCIACAO';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();

        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.emp_empresa     = 1;
        tpEnvio.cli_codigo      = Integer.valueOf(user.Contact.Account.CodigoClienteERP__c);
        tpEnvio.ind_underline   = true;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.method = 'POST';
        parametro.classMethod = 'getNegociacao';
        parametro.className = 'B2BAgrosysNegociacao';
        parametro.customerCode = user.Contact.Account.CodigoClienteERP__c;
        parametro.integrationName = 'Negociação';
        parametro.accountId = user.Contact.AccountId;

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        if( retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0].msg_codigo == 1) {
            return retorno.tp_retorno[0].msg_retorno[0].tp_clineg[0].temnegociacao;
        } else {
            return null;
        }
    }
}