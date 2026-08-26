public without sharing class B2BAgrosysCashback  extends B2BAgrosys {

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer cli_codigo;
        public Integer Hvr_tpverba;
    }
    
    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public B2BAgrosysCashback.CashbackReturn getCashback(String codigoCliente) {
        String cartId;
        String accId;
        
        if(!Test.isRunningTest()) {
			ConnectApi.CartSummary cartSum = ConnectApi.Commercecart.getCartSummary(getStore().Id, null, 'current');
            cartId = cartSum.cartId;
            accId = cartSum.AccountId;
        } else {
            WebCart cartSum = [
                SELECT Id, AccountId
                FROM WebCart
                WHERE WebStoreId = :getStore().Id
                LIMIT 1
            ];
            cartId = cartSum.Id;
            accId = cartSum.AccountId;
        }

        List<Account> accs = [SELECT Id, Unidade__c FROM Account WHERE CodigoClienteERP__c = :codigoCliente OR Id = :accId LIMIT 2];
        Account targetAccount = accs[0];
        if(accs.size() > 1 && targetAccount.CodigoClienteERP__c != codigoCliente){
            targetAccount = accs[1];
        }

        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-CASHBACK';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();

        

        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.emp_empresa     = 1;
        tpEnvio.uni_unidade     = targetAccount.Unidade__c != null ? Integer.valueOf(targetAccount.Unidade__c) : 1;
        tpEnvio.cli_codigo      = Integer.valueOf(codigoCliente);
        tpEnvio.cli_ecommerce   = true;
        tpEnvio.ind_underline   = true;
        tpEnvio.Hvr_tpverba     = 5; // Fixed value 5
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.method = 'POST';
        parametro.classMethod = 'getCashback';
        parametro.className = 'B2BAgrosysCashback';
        parametro.customerCode = codigoCliente;
        parametro.integrationName = 'Cashback';
        parametro.cartId = cartId;
        parametro.accountId = accId;

        CashbackReturn cashReturn = new CashbackReturn();
        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;

        if(retorno.tp_retorno == null || retorno.hasError) {
            cashReturn.hasError = true;
            cashReturn.errorMessage = retorno.errorMessage != null ? retorno.errorMessage : null;
            cashReturn.logId = retorno.logId != null ? retorno.logId : null;
            return cashReturn;
        }

        if(!retorno.tp_retorno.isEmpty() && retorno.tp_retorno[0].msg_codigo == 1) {
            if(retorno.tp_retorno[0].msg_retorno[0].tp_extverba != null) {
                cashReturn.cashbackBalance = String.valueOf(retorno.tp_retorno[0].msg_retorno[0].tp_extverba[0].hvr_saldo).replace('.', ',');
            }
        }
        else {
            if(retorno.tp_retorno[0].msg_codigo != 1) {
                cashReturn.hasError = true;
                cashReturn.errorMessage = getErrorAgrosys(retorno.tp_retorno[0].msg_codigo);
            }
        }
        return cashReturn;
    }

    private static WebStore getStore() {
        WebStore store;
        if (!Test.isRunningTest()){
            Network storeNetWork = [SELECT Name FROM Network WHERE Id =: Network.getNetworkId()];
            store = [SELECT Id FROM WebStore WHERE Name = :storeNetWork.Name];
        } else {
            store = [SELECT Id,Name FROM WebStore LIMIT 1];
        }
        return store;
    }

    public class CashbackReturn {
        public String cashbackBalance;
        public Boolean hasError;
        public String errorMessage;
        public String logId;

        public CashbackReturn() {
            this.hasError = false;
        }
    }
}