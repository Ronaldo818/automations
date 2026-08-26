@SuppressWarnings('PMD.NcssMethodCount,PMD.CognitiveComplexity,PMD.CyclomaticComplexity')
public without sharing class B2BLWRCheckoutControler {
    
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
    /*
    @AuraEnabled
    public static String createOrder(String valorCashback) {
        List<sObject> lstUpdateObj = new List<sObject>();
        valorCashback =  String.isNotBlank(valorCashback) ? valorCashback.replace(',', '.') : null;
        Retorno result = new Retorno();
        Account acc;

        try{
            if(getContactPointAddress()) {
                if(Test.isRunningTest()){
                    result.webCart = [
                        SELECT Id, OrderId__c, AccountId
                        FROM WebCart
                        LIMIT 1
                    ];

                    acc = [
                        SELECT Id, CodigoClienteERP__c
                        FROM Account
                        WHERE Id = :result.webCart.AccountId
                    ];
                }
                else {
                    ConnectApi.CartSummary cartSum = ConnectApi.Commercecart.getCartSummary(getStore().Id, null, 'current');
                    result.webCart = [
                        SELECT Id, OrderId__c, AccountId
                        FROM WebCart
                        WHERE Id = :cartSum.cartId
                    ];

                    acc = [
                        SELECT Id, CodigoClienteERP__c
                        FROM Account
                        WHERE Id = :result.webCart.AccountId
                    ];
                }
               
                List<CartCheckoutSession> lstCartSession = [
                    SELECT Id
                    FROM CartCheckoutSession
                    WHERE WebCartId = :result.webCart.Id
                    AND IsDeleted = false
                    AND State = 'Start'
                ];
                result.CartCheckoutSessionId = lstCartSession[0].Id;

                B2BLWRCreateOrder creator = new B2BLWRCreateOrder(lstCartSession[0].Id, result.webCart);
                String newOrderId = creator.cartToOrder();

                Order ord = [
                    SELECT Id, OrderNumber, Status
                    From Order
                    WHERE Id = :newOrderId
                ];
                
                result.orderId = (String) ord.Id;
                result.orderNumber = ord.OrderNumber;
                // O Pedido é ativado aqui
                // ord.Status = 'Activated';
                ord.ValorCashback__c =  Decimal.valueOf(valorCashback);
                lstUpdateObj.add(ord);
                result.webCart.OrderId__c = ord.Id;
                lstUpdateObj.add(result.webCart);
                update lstUpdateObj;
                //createOrderSummary(ord.Id);
            }
            else {
                AuraHandledException newException = new AuraHandledException('O cliente não possui um endereço de entrega (ContactPointAddress).');
                newException.setMessage('O cliente não possui um endereço de entrega (ContactPointAddress).');
                throw newException; 
            }
        }
        catch(Exception ex){
            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c 			= acc != null ? 'Erro ao tentar criar Pedido  - ' + acc.CodigoClienteERP__c : 'Erro ao tentar criar Pedido';
            log.AccountId__c 		= acc != null ? acc.Id : null;
			log.ApexClass__c 		= 'B2BLWRCheckoutControler';
            log.ApexMethod__c 		= 'execute';
			log.OperationStatus__c	= 'Error';
            log.MessageError__c 	= 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
			insert log;

            result.errorMessage = 'Ocorreu um erro ao tentar finalizar a compra. Informe ao seu representante de vendas ou ao administrador do sistema.';
            result.logId = log.Id;
            result.error = true;
            return System.JSON.serialize(result);
        }

        result.errorMessage = null;
        result.error = false;
        return System.JSON.serialize(result);
    }
    */
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
            cart = [SELECT Id, AccountId, GrandTotalAmount, ValorCashback__c, InitialOrderReferenceNumber FROM WebCart WHERE Id = :parameter.cartId LIMIT 1];

            account = [
                SELECT Id, Name, ShippingStreet, ShippingPostalCode, ShippingCity, ShippingState, ShippingCountry, ComplementoEntrega__c, NumeroEntrega__c, BairroEntrega__c, CNPJ__c, CodigoClienteERP__c, PaisEnderecoEntrega__c, RuaEntrega__c, CEPEntrega__c, CidadeEntrega__r.Name, CidadeEntrega__r.UF__c
                FROM Account
                WHERE Id = :cart.AccountId
            ];

            if(parameter.formaPagamento == 'boleto'){
                cart.PoNumber = cart.InitialOrderReferenceNumber;
                cart.MetodoPagamento__c = 'boleto';
                update cart;
                Retorno retorno = new Retorno();
                retorno.error = false;
                retorno.link = '&boleto=\'\'';
                return JSON.serialize(retorno, true);
            }

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
                /*
                if(parameter.formaPagamento == 'boleto'){
                    for(B2BPagSeguroRetorno.Links link: payment.charges[0].links) {
                        retorno.link = link.href;    
                    }    
                }
                */
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
        }catch(Exception ex){

            System.debug('Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage());

            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c 			= 'ERRO no pagamento - ' + account.CodigoClienteERP__c;
            log.CartId__c 			= cart != null ? cart.Id : null;
            log.AccountId__c 		= account != null ? account.Id : null;
			log.ApexClass__c 		= 'B2BLWRCheckoutControler';
            log.ApexMethod__c 		= 'pay';
			log.OperationStatus__c	= 'Error';
            log.MessageError__c 	= 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
			insert log;

            Retorno retorno = new Retorno();
            retorno.logId = log.Id;
            retorno.errorMessage = 'Não foi possível concluir o Checkout. Informe ao seu representante de vendas ou ao administrador do sistema.';
            retorno.error = true;
            return System.JSON.serialize(retorno);
        }
    }
    /*
    @AuraEnabled
    public static string closeWebCart(String data) {
        WebCart cart;
        Order ord;
        Retorno retorno = new Retorno();
        try {
            PaymentReturn dados = (PaymentReturn) System.JSON.deserialize(data, PaymentReturn.class);
            ord = [
                SELECT Id, AccountId, Account.Name, Account.CodigoClienteERP__c
                FROM Order
                WHERE Id = :dados.orderId
            ];

            if(Test.isRunningTest()){
                cart = [SELECT Id, Status FROM WebCart LIMIT 1];
                getStore();
            }else{
                ConnectApi.CartSummary cartSum = ConnectApi.Commercecart.getCartSummary(getStore().Id, null, 'current');
                cart = [SELECT Id, Status FROM WebCart   WHERE Id = :cartSum.cartId  LIMIT 1];
            }

            cart.Status = 'Closed';
            if(dados.formaPagamento == 'boleto'){
                update cart;
                retorno.link = '&boleto=\'\'';
                return System.JSON.serialize(retorno);
            }

            B2BPagSeguroRetorno payment = dados.pagSeguroReturn;

            if(Test.isRunningTest() && ord.Account.Name.equalsIgnoreCase('error01')) {
                payment = null;
            }

            if(payment != null){
                if(dados.formaPagamento == 'pix'){
                    for(B2BPagSeguroRetorno.links link : payment.qr_codes[0].links){
                        if(link.media == 'image/png'){
                            try {
                                update cart;
                                retorno.link = '&link='+ EncodingUtil.urlEncode(link.href, 'UTF-8') ;
                                
                                if(Test.isRunningTest() && dados.nomeCartao.equalsIgnoreCase('errortest')) {
                                    AuraHandledException newException = new AuraHandledException('isError');
                                    newException.setMessage('isError');
                                    throw newException;
                        		}
                                
                                return System.JSON.serialize(retorno);
                            } catch (Exception ex) {
                                retorno.error = true;
                                retorno.message = ex.getMessage();
                                return System.JSON.serialize(retorno);    
                            }
                        }
                    }
                }
                else if(dados.formaPagamento == 'cc') {
                    try {
                        update cart;
                        retorno.link = '';
                        
                        if(Test.isRunningTest() && dados.nomeCartao.equalsIgnoreCase('errortest')) {
                            AuraHandledException newException = new AuraHandledException('isError');
                            newException.setMessage('isError');
                            throw newException;   
                        }
                        
                        return System.JSON.serialize(retorno);    
                    } catch (Exception ex) {
                        retorno.error = true;
                        retorno.message = ex.getMessage();
                        return System.JSON.serialize(retorno);     
                    }
                }
            }
            
            AuraHandledException newException = new AuraHandledException('Variável "payment" NULL.');
            newException.setMessage('Variável "payment" NULL.');
            throw newException;

        }
        catch(Exception ex) {
            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c 			= ord != null && ord.Account.CodigoClienteERP__c != null ? 'Erro ao tentar finalizar checkout  - ' + ord.Account.CodigoClienteERP__c : 'Erro ao tentar finalizar checkout';
            log.AccountId__c 		= ord != null && ord.AccountId != null ? ord.AccountId : null;
            log.OrderId__c          = ord != null ? ord.Id : null;
			log.ApexClass__c 		= 'B2BLWRCheckoutControler';
            log.ApexMethod__c 		= 'closeWebCart';
			log.OperationStatus__c	= 'Error';
            log.MessageError__c 	= 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
			insert log;

            retorno.errorMessage = 'Não foi possível finalizar a compra. Informe ao seu representante de vendas ou ao administrador do sistema.';
            retorno.logId = log.Id;
            retorno.error = true;
            
            return System.JSON.serialize(retorno);
        }
    }
    */
   /*
    public static void createOrderSummary(String orderId){
        try {
            ConnectApi.OrderSummaryInputRepresentation oderSummaryInput = new ConnectApi.OrderSummaryInputRepresentation();
            oderSummaryInput.orderId = orderId;
            oderSummaryInput.orderLifeCycleType = 'UNMANAGED';
            if (!Test.isRunningTest()) {
                ConnectApi.OrderSummaryCreation.createOrderSummary(oderSummaryInput);
            }
        }
        catch(Exception ex) {
            AuraHandledException newException = new AuraHandledException(ex.getMessage());
        	newException.setMessage(ex.getMessage());
        	throw newException;
        }
    }
    */
   /*
    private static Boolean getContactPointAddress() {
        Account acc;
        User user = [
            SELECT Id, ContactId, Contact.AccountId
            FROM User
            WHERE Id = :UserInfo.getUserId()
        ];

        if(Test.isRunningTest()) {
            acc = [
                SELECT Id, Name, ShippingStreet, ShippingCity, ShippingState, ShippingPostalCode, ShippingCountry, Unidade__c
                FROM Account
                WHERE name = 'isTest'
            ];
        }
        else {
            acc = [
                SELECT Id, Name, ShippingStreet, ShippingCity, ShippingState, ShippingPostalCode, ShippingCountry, Unidade__c
				FROM Account
                WHERE Id = :user.Contact.AccountId
            ];
        }
        
        List<ContactPointAddress> lstCPA = [
            SELECT Id, Name, Street, City, State, PostalCode, Country, AddressType, ParentId, UsageType, IsDefault 
           	FROM ContactPointAddress
			WHERE ParentId = :acc.Id
            AND IsDefault = true
            AND AddressType = 'Shipping'
            Limit 1
        ];
        
        if(lstCPA.isEmpty()){
            return false;
        }else{
            return true;
        }
    }
    */
   /*
    private static WebStore getStore() {
        WebStore store;
        Network storeNetWork;

        if (!Test.isRunningTest()){
            storeNetWork = [SELECT Name FROM Network WHERE Id =: Network.getNetworkId()];
            store = [SELECT Id FROM WebStore WHERE Name =: storeNetWork.Name];
        } else {
            store = [SELECT Id,Name FROM WebStore LIMIT 1];
        }

        return store;
    }
        */
    
}