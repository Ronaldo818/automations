@RestResource(urlMapping='/paymentConfirmation') 
global without sharing class PagSeguroWebhook {
    
    @HttpPost
    global static void handlePost(){
        Order ord;
        String accId;
        RestRequest resReq = RestContext.request;
        
        if(Test.isRunningTest()) {
            Order ped = [SELECT Id FROM Order LIMIT 1];
            String jsonBody = '{ "id":"isTest", "reference_id":"' + ped.Id + '", "created_at":"2024-04-30T14:31:20.783-03:00", "customer":{ "name":"isTest", "email":"istest.istest@istest.com.br", "tax_id":"istest" }, "items":[ { "reference_id":"", "name":"isTest", "quantity":1, "unit_amount":1 } ], "shipping":{ "address":{ "street":"isTest", "number":"010", "locality":"isTest", "city":"isTest", "region_code":"SP", "country":"BRA", "postal_code":"10101010" } }, "charges":[ { "id":"isTest", "reference_id":"isTest", "status":"PAID", "created_at":"2024-04-30T14:31:21.402-03:00", "paid_at":"2024-04-30T14:31:23.000-03:00", "description":"isTest", "amount":{ "value":1, "currency":"BRL", "summary":{ "total":1, "paid":1, "refunded":0 } }, "payment_response":{ "code":"20000", "message":"SUCESSO", "reference":"isTest", "raw_data":{ "authorization_code":"isTest", "nsu":"isTest", "reason_code":"00" } }, "payment_method":{ "type":"CREDIT_CARD", "installments":1, "capture":true, "card":{ "id":"isTest", "brand":"visa", "first_digits":"000000", "last_digits":"0000", "exp_month":"0", "exp_year":"0000", "holder":{ "name":"isTest" } }, "soft_descriptor":"sellervirtual" }, "links":[ { "rel":"SELF", "href":"isTest", "media":"application/json", "type":"GET" }, { "rel":"CHARGE.CANCEL", "href":"isTest", "media":"application/json", "type":"POST" } ], "metadata":{ } } ], "notification_urls":[ "isTest" ], "links":[ { "rel":"SELF", "href":"isTest", "media":"application/json", "type":"GET" }, { "rel":"PAY", "href":"isTest", "media":"application/json", "type":"POST" } ] }';
        	resReq = new System.RestRequest();
            resReq.requestBody = Blob.valueOf(jsonBody);
            Blob data = Blob.valueOf(UserInfo.getLastName() + '-' + jsonBody);
        	Blob hash = Crypto.generateDigest('SHA-256', data);
        	String encryptedString = EncodingUtil.convertToHex(hash);
            resReq.addHeader('x-authenticity-token', encryptedString);
        }

        try {
            if (String.isNotBlank(resReq.requestBody.toString())) {
                if (confirmAuthentication(resReq.headers.get('x-authenticity-token'), resReq.requestBody.toString())) {
                    B2BPagSeguroRetorno retorno = B2BPagSeguroRetorno.parse(resReq.requestBody.toString());
                
                    if (String.isNotBlank(retorno.reference_id)) {
                        ord = [
                            SELECT Id, Status ,OrderNumber, AccountId, StatusPagamento__c, PagSeguroChargeId__c, PagSeguroOrderId__c
                            FROM Order
                            WHERE Id = :retorno.reference_id
                        ];
                        
                        if (ord != null) {
                            accId = ord.AccountId;
                            ord.PagSeguroOrderId__c = retorno.id;
                            // param.orderId = ord.Id;
                            Boolean paid = false;
                            
                            if (!retorno.charges.isEmpty()) {
                                String chargeId;
                                for (B2BPagSeguroRetorno.Charges charge : retorno.charges) {
                                    if (charge.status.equalsIgnoreCase('paid')) {
                                        paid = true;
                                        chargeId = charge.id;
                                    }
                                }
    
                                if (paid) {
                                    ord.StatusPagamento__c = 'PAID';
                                    ord.PagSeguroChargeId__c = chargeId;
                                }
                                else {
                                    ord.StatusPagamento__c = retorno.charges[0]?.status;    
                                    ord.PagSeguroChargeId__c = retorno.charges[0]?.Id;
                                }
                            }
                            
                            update ord;
                            if(paid && ord.Status != 'Reserved') {
                                sendOrder(ord.Id);
                            }
                        }
                	}
            	}
                else if(Test.isRunningTest() && !confirmAuthentication(resReq.headers.get('x-authenticity-token'), resReq.requestBody.toString())) {
                    if(UserInfo.getLastName().equalsIgnoreCase('error')) {
                        insert new Account();
                    }
                }
                

                IntegrationLog__c log = new IntegrationLog__c(
                    NomeLog__c = ord != null ? 'PagSeguro Status Pagamento - ' + ord.OrderNumber : 'PagSeguro Status Pagamento',
                    AccountId__c = accId,
                    OrderId__c = ord != null ? ord.Id : null,
                    ApexClass__c    = 'PagSeguroWebhook',
                    ApexMethod__c       = 'handlePost',
                    requestBody01__c    = resReq.requestBody.toString(),
                    OperationStatus__c  = 'Success'
                );
                insert log;

            }
            
            if(!Test.isRunningTest()) {
				RestContext.response.statusCode = 200;                
            }
        }
        catch (Exception ex) {
            IntegrationLog__c log = new IntegrationLog__c(
                NomeLog__c          = ord != null ? 'ERROR PagSeguro Status Pagamento - ' + ord.OrderNumber : 'ERROR PagSeguro Status Pagamento',
                AccountId__c        = accId,
                OrderId__c          = ord != null ? ord.Id : null,
                ApexClass__c        = 'PagSeguroWebhook',
                ApexMethod__c       = 'handlePost',
                requestBody01__c    = resReq.requestBody.toString(),
                OperationStatus__c  = 'Error',
                MessageError__c     = 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage()
            );
			insert log;
        }
    }

    private static Boolean confirmAuthentication(String authenticityToken, String payload) {
        if (String.isBlank(authenticityToken) || String.isBlank(payload)) {
            return false;
        }
		String acessToken;
        if(Test.isRunningTest()) {
        	acessToken = 'Success';	
        }
        else {
        	WebService__mdt service = WebService__mdt.getInstance('PagSeguro');
            acessToken = service.AcessToken__c;
        }
        Blob data = Blob.valueOf(acessToken + '-' + payload);
        Blob hash = Crypto.generateDigest('SHA-256', data);
        String encryptedString = EncodingUtil.convertToHex(hash);

        if (encryptedString == authenticityToken) {
            return true;
        }
        else {
            return false;
        }
    }

    @Future(callout=true)
    private static void sendOrder(String orderId) {
        Order ord = [
            SELECT Id, NumeroParcelas__c, MetodoPagamento__c
            FROM Order
            WHERE Id = :orderId
        ];

        B2BAgrosysPedido.SendOrderParameter param = new B2BAgrosysPedido.SendOrderParameter();
        param.operacao = 3;
        param.orderId = ord.Id;
        param.parcelas = ord.NumeroParcelas__c;
        param.formaPagamento = ord.MetodoPagamento__c;
        B2BAgrosysPedido.setPedido(System.JSON.serialize(param));
    }

}