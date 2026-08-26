public with sharing class B2BPagSeguro {

    public class ApexInput {
        public String cartId;
        public String formaPagamento;
        public String cartaoCriptografado;
        public String nomeCartao;
        public String numeroCartao;
        public String mesCartao;
        public String anoCartao;
        public String cvvCartao;
        public String parcelas;
        public WebCart cart;
        public Account account;

        public ApexInput(B2BLWRCheckoutControler.PaymentParameter parameter) {
            this.formaPagamento = parameter.formaPagamento;
            this.nomeCartao = parameter.nomeCartao;
            this.cartaoCriptografado = parameter.encryptedCard;
            this.numeroCartao = parameter.numeroCartao;
            this.mesCartao = parameter.mesCartao;
            this.anoCartao = parameter.anoCartao;
            this.cvvCartao = parameter.cvvCartao;
            this.parcelas = parameter.parcelas;
        }
    }

    public static B2BPagSeguroRetorno executar(ApexInput input) {
        WebCart cart = input.cart;
        Account account = input.account;

        String json = B2BPagSeguroRequisicao.parse(input, cart);
        B2BPagSeguroRetorno b2bResposta = B2BPagSeguro.postRequest(json, cart.Id, account);

        if(b2bResposta == null || String.isBlank(b2bResposta.id)) {
            AuraHandledException newException = new AuraHandledException('Erro ao tentar realizar o pagamento. Informe ao administrador do sistema o ID do Log: ' + b2bResposta.logId);
        	newException.setMessage('Erro ao tentar realizar o pagamento. Informe ao administrador do sistema o ID do Log: ' + b2bResposta.logId);
        	throw newException;
        }
        else {
            cart.PoNumber = b2bResposta.reference_id;
            if(input.formaPagamento.equalsIgnoreCase('cc')) {
                cart.MetodoPagamento__c = 'cc';
                cart.PagSeguroPaymentReferenceNumber__c = b2bResposta.charges[0]?.payment_response.reference;
                cart.BandeiraCartao__c = b2bResposta.charges[0].payment_method.card.brand;
                cart.NumeroParcelas__c = b2bResposta.charges[0].payment_method.installments.toString() + 'x';
            } else if(input.formaPagamento.equalsIgnoreCase('pix')) {
                cart.MetodoPagamento__c = 'pix';
                cart.ChavePix__c = b2bResposta.qr_codes[0]?.text;
                cart.PaymentExpiration__c = b2bResposta.qr_codes[0].expiration_date;
                for(B2BPagSeguroRetorno.links link: b2bResposta.qr_codes[0].links) {
                    if(link.media == 'image/png'){
                        cart.URLPix__c = link.href;
                        break;
                    }
                }
            } else if(input.formaPagamento.equalsIgnoreCase('boleto')) {
                cart.MetodoPagamento__c = 'boleto';
            }
            update cart;
            return b2bResposta;
        }
    }

    static B2BPagSeguroRetorno postRequest(String json, String cartId, Account account) {
        System.debug('postRequest');
        WebService__mdt service = WebService__mdt.getInstance('PagSeguro');
        Http http = new Http();
        HttpRequest request = new HttpRequest();
        HttpResponse response = new HttpResponse();
        request.setHeader('Content-type', 'application/json');
        // request.setHeader('Authorization', service.Password__c);
        request.setHeader('Authorization', 'Bearer ' + service.AcessToken__c);
        request.setEndpoint(service.BaseUrl__c+service.Directory__c);
        request.setBody(json);
        request.setMethod('POST');
        response = http.send(request);
        
        IntegrationLog__c log = new IntegrationLog__c();
        log.NomeLog__c = 'PagSeguro - ' + account.CodigoClienteERP__c;
		//log.OrderId__c = String.isNotBlank(orderId) ? orderId : null;
        log.CartId__c = cartId;
        log.AccountId__c = String.isNotBlank(account.Id) ? account.Id : null;
		log.ApexClass__c = 'B2BPagSeguro';
        log.ApexMethod__c = 'postRequest';
		log.EndPoint__c = request.getEndpoint();
		log.IntegrationMethod__c = 'POST';
		log.IntegrationStatusCode__c = response.getStatusCode()+' : '+response.getStatus();
		log.RequestBody01__c = request.getBody();
        log.responseBody01__c =  response.getBody();
		log.OperationStatus__c = response.getStatusCode() != 200 && response.getStatusCode() != 201 ? 'Error' : 'Success';
		insert log;

        if(response.getStatusCode() == 200 || response.getStatusCode() == 201) {
            return B2BPagSeguroRetorno.parse(response.getBody());
        }
        else {
            if(Test.isRunningTest() && account.Name.equalsIgnoreCase('error')) {
                AuraHandledException newException = new AuraHandledException('isError');
				newException.setMessage('isError');
				throw newException;
            }
            B2BPagSeguroRetorno retorno = new B2BPagSeguroRetorno();
            retorno.logId = log.Id;
            return retorno;
        }
    }
}