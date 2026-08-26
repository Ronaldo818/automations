public with sharing class B2BAgrosysTabelaPreco extends B2BAgrosys {

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer pre_tiporeg;
        public Integer pit_tabpre;
        public Integer pro_codigo;
        public Integer cli_codigo;
    }
    
    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public B2BAgrosysRetorno getCustomerPricebook2(Account account) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-TBPRECO ';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();

        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_codigo      = Integer.valueOf(account.CodigoClienteERP__c);
        tpEnvio.uni_unidade     = Integer.valueOf(account.Unidade__c);
        tpEnvio.emp_empresa     = 1;
        tpEnvio.cli_ecommerce   = true;
        tpEnvio.ind_underline   = true;
        tpEnvio.pre_tiporeg     = 0;
        tpEnvio.pit_tabpre      = 0;
        tpEnvio.pro_codigo      = 0;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body              = jsonBody;
        parametro.integrationName   = 'Lista de Preços';
        parametro.customerCode      = account.CodigoClienteERP__c;
        parametro.method            = 'POST';
        parametro.classMethod       = 'getCustomerPricebook2';
        parametro.className         = 'B2BAgrosysTabelaPreco - ' + account.CodigoClienteERP__c;
        parametro.accountId         = account.Id;

        System.debug('JSONBODY =>');
        System.debug(jsonBody);
        return String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
    }
}