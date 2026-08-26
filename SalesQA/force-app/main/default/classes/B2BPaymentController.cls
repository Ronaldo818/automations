public without sharing class B2BPaymentController {
    
    public static final String CHECKOUT_ERROR_MESSAGE = 'Não foi possível concluir o Checkout. Informe ao seu representante de vendas ou ao administrador do sistema.';
    
    public class Retorno{
        public PaymentReturn paymentReturn;
        public String message;
        public String orderId;
        public String orderNumber;
        public WebCart webCart;
        public String paymentStatus;
        public String boletoQrcod;
        public String CartCheckoutSessionId;
        public String cashbackBalance;
        public String link;
        public Boolean error = false;
        public String logId;
        public String errorMessage;
    }
    
    public class PaymentReturn {
        public B2BPagSeguroRetorno pagSeguroReturn;
        public String cartId;
        public String parcelas;
        public String cvvCartao;
        public String anoCartao;
        public String mesCartao;
        public String numeroCartao;
        public String nomeCartao;
        public String formaPagamento;
        public String valorCashback;
    }
    
    public class PaymentParameter {
        public String cartId;
        public String encryptedCard;
        public String parcelas;
        public String cvvCartao;
        public String anoCartao;
        public String mesCartao;
        public String numeroCartao;
        public String nomeCartao;
        public String formaPagamento;
        public String valorCashback;
        public Integer operacao;
    }

    @AuraEnabled
    public static String setCashBackValue(Decimal valorCashback) {
    List<Cashback__c> cashback = [SELECT Saldo__c FROM Cashback__c WHERE Cliente__c = :CommerceUtil.getCurrentEffectiveAccountId()];
        if(valorCashback != null && !cashback.isEmpty() ){
            if(valorCashback !=0 && valorCashback > cashback[0].Saldo__c){
                Retorno retorno = new Retorno();
                retorno.errorMessage = 'Uso de cashback inválido.';
                retorno.error = true;
                return System.JSON.serialize(retorno);
            }
            else{
                WebCart c = new WebCart();
                c.Id = CommerceUtil.getCartSummary().cartId;
                c.ValorCashback__c = valorCashback;
                update c;
                return null;
            }
        }
        return null;
    }
    @AuraEnabled
    public static String getCashback() {
        Retorno retorno = new Retorno();
        // B2BAgrosysCashback.CashbackReturn cashbackReturn = new B2BAgrosysCashback.CashbackReturn();
        try {
                
                List<Cashback__c> cb = [ SELECT Saldo__c FROM Cashback__c WHERE Cliente__c = :CommerceUtil.getCurrentEffectiveAccountId() LIMIT 1];
                if(cb.isEmpty()){
                    retorno.cashbackBalance = '0.0';
                }
                else {
                    retorno.cashbackBalance = String.valueOf(cb[0].Saldo__c);
                }
                return System.JSON.serialize(retorno);
        } catch (Exception ex) {
            return System.JSON.serialize('');
        }
    }
    /*
    @AuraEnabled
    public static Boolean checkCard(String jsonParameter){
        return (Integer)CommerceCieloIntegration.verifyCard((Map<String,Object>)JSON.deserializeUntyped(jsonParameter)).get('Status') == 1;
    }
    */
    @AuraEnabled
    public static String pay(String jsonParameter) {
		WebCart cart;
        Account account;

        System.debug('PAY jsonParameter ==> ' + jsonParameter);
        B2BLWRCheckoutControler.PaymentParameter parameter = (B2BLWRCheckoutControler.PaymentParameter) System.JSON.deserialize(jsonParameter, B2BLWRCheckoutControler.PaymentParameter.class);
        
        try{
        /* 
            order = [
                SELECT Id, OrderNumber, AccountId, totalAmount, ContactId__c, ValorCashback__c, GrandTotalAmount
                FROM Order
                WHERE Id = :parameter.orderId
            ];
        */  
            cart = [SELECT Id, AccountId, GrandTotalAmount, ValorCashback__c, InitialOrderReferenceNumber, PaymentMethodId, CurrencyIsoCode FROM WebCart WHERE Id = :parameter.cartId LIMIT 1];
            if(cart.ValorCashback__c == null){
                cart.ValorCashback__c = 0;
            }
            account = [
                SELECT Id, Email__c, cleanCpfCnpj__c, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
                FROM Account
                WHERE Id = :cart.AccountId
            ];

            CartDeliveryGroup cdg = [SELECT Id, DeliverToCity, DeliverToCountry, DeliverToState, DeliverToStreet, DeliverToPostalCode FROM CartDeliveryGroup WHERE CartId = :cart.Id LIMIT 1];

            switch on parameter.formaPagamento {
                when 'boleto' {
                    cart.PoNumber = cart.InitialOrderReferenceNumber;
                    cart.MetodoPagamento__c = 'boleto';
                    update cart;
                    Retorno retorno = new Retorno();
                    retorno.error = false;
                    retorno.link = '&paymentType=boleto&boleto=\'\'';
                    return JSON.serialize(retorno, true);
                }
                when 'cc' {
                    cart.MetodoPagamento__c = 'cc';
                    Retorno retorno = new Retorno();
                    Map<String,Object> data = (Map<String,Object>)JSON.deserializeUntyped(jsonParameter);

                    CommerceCieloIntegration verifyIntegration = new CommerceCieloIntegration('CieloVerifyCard');
                    //verifyIntegration.inputData = (Map<String,Object>)data.get('cardParameter');
                    Map<String,Object> cardData = (Map<String,Object>)data.get('cardParameter');
                    verifyIntegration.cardInput = dataToCieloCard(cardData);
                    verifyIntegration.createPayload();
                    verifyIntegration.sendPayload();

                    if(!verifyIntegration.logs.isEmpty()){
                        verifyIntegration.processFinalDMLs();
                        Retorno ret = new Retorno();
                        ret.logId = verifyIntegration.logs[0].Id;
                        ret.errorMessage = CHECKOUT_ERROR_MESSAGE;
                        ret.error = true;
                        return System.JSON.serialize(ret);
                    }

                    Boolean isValid = (Boolean)verifyIntegration.responseData.get('Valid');
                    if(isValid == null){
                        // v1 || v2
                        isValid = !(((!verifyIntegration.responseData.containsKey('Status') && (Integer)verifyIntegration.responseData.get('Status') != 1)) || (!verifyIntegration.responseData.containsKey('Valid') && (Boolean)verifyIntegration.responseData.get('Valid') == false));
                    }

                    if(!isValid){
                        retorno.error = true;
                        // v2
                        retorno.errorMessage = (String)verifyIntegration.responseData.get('ReturnMessage');
                        // v1
                        if(String.isBlank(retorno.errorMessage)) retorno.errorMessage = (String)verifyIntegration.responseData.get('ProviderReturnMessage');
                        verifyIntegration.processFinalDMLs();
                        return JSON.serialize(retorno, true);
                    }

                    CommerceCieloIntegration integration = new CommerceCieloIntegration('CieloCreatePayment');
                    //integration.inputData = data;
                    integration.cardInput = verifyIntegration.cardInput;
                    integration.installments = Integer.valueOf(String.valueOf((String)cardData.get('parcelas')).replace('x',''));
                    integration.cart = cart;
                    integration.acc = account;
                    integration.cdg = cdg;
                    integration.process();

                    if(!integration.logs.isEmpty()){
                        integration.processFinalDMLs();
                        Retorno ret = new Retorno();
                        ret.logId = integration.logs[0].Id;
                        ret.errorMessage = CHECKOUT_ERROR_MESSAGE;
                        ret.error = true;
                        return System.JSON.serialize(ret);
                    }

                    cart.NumeroParcelas__c = parameter.parcelas;
                    update cart;

                    verifyIntegration.processFinalDMLs();
                    
                    Map<String,Object> paymentData = (Map<String,Object>)integration.responseData.get('Payment');
                    Integer cieloStatus = (Integer)paymentData.get('Status');
                    retorno.error = cieloStatus != 2 && cieloStatus != 12 && cieloStatus != 1;
                    retorno.errorMessage = (String)paymentData.get('ReturnMessage');
                    retorno.link = '&paymentType=cc';
                    return JSON.serialize(retorno, true);
                }
                when 'pix' {
                    cart.MetodoPagamento__c = 'pix';
                    CommerceCieloIntegration integration = new CommerceCieloIntegration('CieloCreatePayment');
                    integration.cart = cart;
                    integration.acc = account;
                    integration.cdg = cdg;
                    integration.process();

                    if(!integration.logs.isEmpty()){
                        Retorno ret = new Retorno();
                        ret.logId = integration.logs[0].Id;
                        ret.errorMessage = CHECKOUT_ERROR_MESSAGE;
                        ret.error = true;
                        return System.JSON.serialize(ret);
                    }

                    update cart;

                    Retorno retorno = new Retorno();
                    Map<String,Object> paymentData = (Map<String,Object>)integration.responseData.get('Payment');
                    Integer cieloStatus = (Integer)paymentData.get('Status');
                    retorno.error = cieloStatus != 2 && cieloStatus != 12 && cieloStatus != 1;
                    retorno.errorMessage = (String)paymentData.get('ReturnMessage');
                    if(String.isBlank(retorno.errorMessage)) retorno.errorMessage = (String)paymentData.get('ProviderReturnMessage');
                    retorno.link = '&paymentType=pix';
                    return JSON.serialize(retorno, true);
                }
                when else {
                    Retorno retorno = new Retorno();
                    retorno.error = true;
                    retorno.errorMessage = 'Forma de Pagamento não especificada ou inválida: ' + parameter.formaPagamento;
                    return JSON.serialize(retorno, true);
                }
            }

            /*

            B2BPagSeguro.ApexInput input = new B2BPagSeguro.ApexInput(parameter);
            input.account = account;
            input.cart = cart;

            B2BPagSeguroRetorno payment = B2BPagSeguro.executar(input);
            
            PaymentReturn paymentReturn = new PaymentReturn();
            paymentReturn.anoCartao         = parameter.anoCartao;
            paymentReturn.cvvCartao         = parameter.cvvCartao;
            paymentReturn.formaPagamento    = parameter.formaPagamento;
            paymentReturn.mesCartao         = parameter.mesCartao;
            paymentReturn.nomeCartao        = parameter.nomeCartao;
            paymentReturn.numeroCartao      = parameter.numeroCartao;
            paymentReturn.cartId           = parameter.cartId;
            paymentReturn.pagSeguroReturn   = payment;
            paymentReturn.parcelas          = parameter.parcelas;
            paymentReturn.valorCashback     = parameter.valorCashback;

            Retorno retorno = new Retorno();
            retorno.paymentReturn = paymentReturn;

            if(payment != null){
                if(payment.charges != null && !payment.charges.isEmpty()) {
                    if(payment.charges[0].status.equalsIgnoreCase('declined')) {
                        retorno.error = true;
                        retorno.errorMessage = payment.charges[0].payment_response.message;
                    }
                }
                
                //if(parameter.formaPagamento == 'boleto'){
                //    for(B2BPagSeguroRetorno.Links link: payment.charges[0].links) {
                //        retorno.link = link.href;    
                //    }    
                //}
                
                else if(parameter.formaPagamento == 'pix') {
                    for(B2BPagSeguroRetorno.links link: payment.qr_codes[0].links) {
                        if(link.media == 'image/png'){
                            retorno.link = '&link='+ EncodingUtil.urlEncode(link.href, 'UTF-8');
                            break;
                        }
                    }
                }
            }
            
            if(Test.isRunningTest() && account.Name.equalsIgnoreCase('error01')) {
            	AuraHandledException newException = new AuraHandledException('isError');
                newException.setMessage('isError');
                throw newException;
            }
            
            return System.JSON.serialize(retorno);
        */
        }catch(Exception ex){

            System.debug('Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage());
            /*
            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c 			= 'ERRO no pagamento - ' + account.CodigoClienteERP__c;
            log.CartId__c 			= cart != null ? cart.Id : null;
            log.AccountId__c 		= account != null ? account.Id : null;
			log.ApexClass__c 		= 'B2BLWRCheckoutControler';
            log.ApexMethod__c 		= 'pay';
			log.OperationStatus__c	= 'Error';
            log.MessageError__c 	= 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
			insert log;
            */

            LogErro__c log = new LogErro__c(TipoAutomacao__c='Apex',NomeAutomacao__c = 'B2BPaymentController.pay',Status__c = 'Erro',Nome__c = 'Log de Erro de Pagamento Cielo - Interface B2B X Integração Cielo', TextoErro__c = ex.getMessage() + '\n' + ex.getStackTraceString());
            insert log;

            Retorno retorno = new Retorno();
            retorno.logId = log.Id;
            retorno.errorMessage = CHECKOUT_ERROR_MESSAGE;
            retorno.error = true;
            return System.JSON.serialize(retorno);
        }  
    }
    
    public static CommerceCieloIntegration.Card dataToCieloCard(Map<String,Object> data){
        CommerceCieloIntegration.Card c = new CommerceCieloIntegration.Card();
        c.CardNumber = (String)data.get('numeroCartao');
        c.ExpirationDate = String.valueOf((String)data.get('mesCartao')).leftPad(2, '0') + '/' + (String)data.get('anoCartao');
        c.Holder = (String)data.get('nomeCartao');
        c.SecurityCode = (String)data.get('cvvCartao');
        c.CardType = 'CreditCard';
        return c;
    }
}