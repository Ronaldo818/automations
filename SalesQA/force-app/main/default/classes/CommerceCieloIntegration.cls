public without sharing class CommerceCieloIntegration {

    HttpRequest request = new HttpRequest();
    HttpResponse response;

    public List<LogErro__c> logs = new List<LogErro__c>();
    public List<IntegrationLog__c> iLogs = new List<IntegrationLog__c>();

    private String operation;
    private String paymentType;
    private String data;
    private PaymentGroup pg;
    private CardPaymentMethod cpm;
    private AlternativePaymentMethod apm;
    private PaymentAuthorization pa;
    //private PaymentAuthAdjustment paa;
    private Boolean cancelPendingPayment = false;
    private ContentVersion cv;

    private Payment p;

    private Order o;

    private Refund r;
    private RefundLinePayment rlp;
    
    
    //public Map<String,Object> inputData = new Map<String,Object>();
    public Card cardInput = new Card();
    public Integer installments = 1;
    public Map<String,Object> responseData = new Map<String,Object>();
    public WebCart cart;
    public Account acc;
    public CartDeliveryGroup cdg;

    public string message;

    public static final String PAID_ORDER_STATUS = 'Activated';
    public static final String CANCELED_ORDER_STATUS = 'Canceled';
    public static final String REFUNDED_ORDER_STATUS = 'Canceled';

    public static final String providerName = System.DomainParser.parse(URL.getOrgDomainUrl()).getSandboxName() != null ? 'Simulado' : 'Cielo30';
    public static final String providerNamePix = System.DomainParser.parse(URL.getOrgDomainUrl()).getSandboxName() != null ? 'Simulado' : 'Cielo2';
    
    public static final Map<String,String> paymentTypeMap = new Map<String,String>{
        'cc' => 'CreditCard',
        'pix' => 'Pix'
    };

    public class CardOnFile{
        public String Usage; //'First' or 'Used'
        public String Reason; // 'Recurring' or 'Unscheduled'
    }

    public class Card{
        public String CardToken;
        public String CardNumber;
        public String Holder;
        public String ExpirationDate; // MM/AAAA
        public String SecurityCode;
        public String Brand;
        public String CardType; // CreditCard/DebitCard
        //public Boolean SaveCard = false;
        //public String Alias = '';
        public CardOnFile CardOnFile;
    }

    public class VerifyCardBody{
        VerifyCardBody(){
            Card = new Card();
            Provider = providerName;
        }
        public String Provider;
        public Card Card;
    }
    
    public class Customer{
        public String Name; // Required
        public String Identity;
        public String IdentityType;
        public String Email;
        public String Birthdate;
        public String IpAddress;
        public Address Address;
        public Address DeliveryAddress;
    }

    public class Address{
        public String Street;
        public String Number_rplc;
        public String Complement;
        public String ZipCode;
        public String City;
        public String State;
        public String Country;
        public String District;
    }

    public class CieloPayment{
        //CieloPayment(){
        //    this.Provider = providerName;
        //}
        public String Provider;
        public String Type;
        public Integer Amount;
        public String Currency_rplc = 'BRL';
        public String Country = 'BRA';
        public Integer Installments = 1;
        public String Interest = 'ByMerchant';
        public Boolean Capture = true;
        public Boolean Authenticate = false;
        public Boolean Recurrent = false;
        public String SoftDescriptor = 'MercadoAvivar'; // Max 13 char
        public Boolean DoSplit = false;
        public Boolean Tip = false;
        public Card CreditCard;
        public Card DebitCard;
        public QrCode QrCode;
        public InitiatedTransactionIndicator InitiatedTransactionIndicator; // Apenas para Mastercard
        public List<ExtraDataCollection> ExtraDataCollection;
    }
    public class QrCode{
        public Integer Expiration; // tempo em segundos

    }

    public class InitiatedTransactionIndicator{
        public String Category; // - “C1”: transação inciada pelo portador do cartão; - “M1”: transação recorrente ou parcelada iniciada pela loja; - “M2”: transação iniciada pela loja. Tamanho: 2.
    }

    public class ExtraDataCollection{
        ExtraDataCollection(String name, String Value){
            this.Name = name;
            this.Value = value;
        }
        String Name;
        String Value;
    }

    public class CreatePaymentBody{
        CreatePaymentBody(){
            Customer = new Customer();
            Payment = new CieloPayment();
        }
        public String MerchantOrderId;
        public Customer Customer;
        public CieloPayment Payment;
    }
    public CommerceCieloIntegration(String operation){
        this(operation, null);    
    }

    public CommerceCieloIntegration(String operation, String data) {
        try{
            this.operation = operation;
            this.data = data;
            Integration__mdt param = [SELECT Endpoint__c, Method__c, Timeout__c FROM Integration__mdt WHERE DeveloperName = :operation LIMIT 1];
            if(operation == 'CieloGetPayment' && String.isNotBlank(this.data)) param.EndPoint__c = param.EndPoint__c.replace('{PaymentId}',data);
            else if(operation == 'CieloRefund'){
                param.EndPoint__c = param.EndPoint__c.replace('{PaymentId}',data)/*.replace('{Amount}', String.valueOf(Integer.valueOf([SELECT Amount FROM Payment WHERE GatewayRefNumber = :data][0].Amount * 100.0)))*/;
            }
            request.setEndpoint(param.Endpoint__c);
            request.setMethod(param.Method__c);
            request.setTimeout(Integer.valueOf(param.Timeout__c));  
            request.setHeader('Content-Type', 'application/json');
        } catch (Exception e) {
            LogErro__c l = new LogErro__c(TipoAutomacao__c='Apex', NomeAutomacao__c = 'CommerceCieloIntegration', Status__c = 'Erro', Nome__c = 'Log de Integração Cielo - Inicialização da Integração',TextoErro__c = e.getMessage() + '\n' + e.getStackTraceString());
            logs.add(l);
        }
    }    

    public void createPayload(){
        try{
            switch on operation {
                when 'CieloVerifyCard' {
                    Card c = cardInput;
                    //c.Brand = 'Visa';//getCardBrand(c.CardNumber.left(6));

                    //b.Provider = provider;
                    //b.Card.Brand = (String)inputData.get('Brand');
                    /*
                    c.CardNumber = (String)inputData.get('numeroCartao');
                    c.ExpirationDate = String.valueOf((String)inputData.get('mesCartao')).leftPad(2, '0') + '/' + (String)inputData.get('anoCartao');
                    c.Holder = (String)inputData.get('nomeCartao');
                    c.SecurityCode = (String)inputData.get('cvvCartao');
                    c.Type = 'CreditCard';
                    */

                    request.setBody(JSON.serialize(c, true).replaceAll('_rplc', ''));
                }
                when 'CieloCreatePayment' {
                    paymentType = paymentTypeMap.get(cart.MetodoPagamento__c);
                    CreatePaymentBody b = new CreatePaymentBody();

                    b.MerchantOrderId = cart.InitialOrderReferenceNumber;

                    b.Customer.Name = acc.Name;
                    b.Customer.Email = acc.Email__c;
                    b.Customer.Identity = acc.cleanCpfCnpj__c;
                    b.Customer.IdentityType = acc.cleanCpfCnpj__c.length() == 11 ? 'CPF' : 'CNPJ';
                    b.Customer.DeliveryAddress = new Address();
                    b.Customer.DeliveryAddress.City = cdg.DeliverToCity;
                    //b.Customer.DeliveryAddress.Complement;
                    b.Customer.DeliveryAddress.Country = cdg.DeliverToCountry;
                    b.Customer.DeliveryAddress.State = cdg.DeliverToState;
                    b.Customer.DeliveryAddress.Street = cdg.DeliverToStreet;
                    b.Customer.DeliveryAddress.ZipCode = cdg.DeliverToPostalCode;
                    b.Customer.Address = b.Customer.DeliveryAddress;

                    b.Payment.Type = paymentType;
                    b.Payment.Amount = Integer.valueOf((cart.GrandTotalAmount - cart.ValorCashback__c) * 100);
                    b.Payment.ExtraDataCollection = new List<ExtraDataCollection>{new ExtraDataCollection('CartId', cart.Id)};
                    switch on paymentType {
                        when 'CreditCard' {
                            //Map<String,Object> cardData = (Map<String,Object>)inputData.get('cardParameter');
                            b.Payment.Installments = installments;
                            b.Payment.CreditCard = cardInput;
                            b.Payment.CreditCard.CardType = null;
                            /*
                            b.Payment.CreditCard.CardNumber = (String)cardData.get('numeroCartao');
                            b.Payment.CreditCard.ExpirationDate = String.valueOf((String)cardData.get('mesCartao')).leftPad(2, '0') + '/' + (String)cardData.get('anoCartao');
                            b.Payment.CreditCard.Holder = (String)cardData.get('nomeCartao');
                            b.Payment.CreditCard.SecurityCode = (String)cardData.get('cvvCartao');
                            //b.Payment.CreditCard.Brand = getCardBrand(b.Payment.CreditCard.CardNumber.left(6));
                            //b.Payment.CreditCard.Type = 'CreditCard';
                            */
                            cart.BandeiraCartao__c = b.Payment.CreditCard.Brand;
                        }
                        when 'Pix' {
                            b.Payment.Provider = providerNamePix;
                            b.Payment.QrCode = new QrCode();
                            b.Payment.QrCode.Expiration = Integer.valueOf([SELECT Value__c FROM StandardValue__mdt WHERE DeveloperName = 'PixHoursToExpire' LIMIT 1][0].Value__c) * 3600;
                        }
                        when else {
                            System.debug('Tipo de Pagamento Não Programado');
                        }
                    }
                    request.setBody(JSON.serialize(b, true).replaceAll('_rplc', ''));
                }
                when 'CieloGetPayment','CieloRefund' {
                    System.debug('No body required');
                }
                when else {
                    System.debug('Invalid Operaton');
                }
            }
        } catch (Exception e) {
            LogErro__c l = new LogErro__c(TipoAutomacao__c='Apex', NomeAutomacao__c = 'CommerceCieloIntegration', Status__c = 'Erro', Nome__c = 'Log de Integração Cielo - Criação da Requisição',TextoErro__c = e.getMessage() + '\n' + e.getStackTraceString() + '\n cardInput:\n' + cardInput + '\n installments:\n' + installments);
            logs.add(l);
        }
    }

    public void sendPayload(){
        IntegrationLog__c ilog = new IntegrationLog__c();
        try {
            ilog.ApexClass__c = 'CommerceCieloIntegration';
            iLog.ApexMethod__c = 'sendPayload';
            iLog.CartId__c = cart != null ? cart.Id : null;
            iLog.AccountId__c = acc != null ? acc.Id : null;
            iLog.EndPoint__c = request.getEndpoint();
            iLog.IntegrationMethod__c = request.getMethod();
            iLog.NomeLog__c = 'Log de Integração Cielo - ' + operation;
            iLog.Request__c = maskSensitiveData(request.getBody().left(32768));
            response = new HTTP().send(request);
            iLog.IntegrationStatusCode__c = String.valueOf(response.getStatusCode());
            iLog.Response__c = response.getBody().left(32768);
            iLogs.add(ilog);
            if(iLog.IntegrationStatusCode__c != '200' && iLog.IntegrationStatusCode__c != '201'){
                
                LogErro__c l = new LogErro__c(TipoAutomacao__c='Apex', NomeAutomacao__c = 'CommerceCieloIntegration', Status__c = 'Erro', Nome__c = 'Log de Integração Cielo - Envio da Requisição',TextoErro__c = 'Falha na integração, verificar logs de integração');
                logs.add(l);
                return;
            }
            responseData = (Map<String,Object>)JSON.deserializeUntyped(response.getBody());
        } catch (Exception e) {
            LogErro__c l = new LogErro__c(TipoAutomacao__c='Apex', NomeAutomacao__c = 'CommerceCieloIntegration', Status__c = 'Erro', Nome__c = 'Log de Integração Cielo - Envio da Requisição',TextoErro__c = e.getMessage() + '\n' + e.getStackTraceString());
            logs.add(l);
        }
    }

    public void processReturn(){
        try{
            switch on operation {
                when 'CieloVerifyCard' {

                }
                when 'CieloCreatePayment' {
                    pg = new PaymentGroup();
                    Map<String,Object> paymentData = (Map<String,Object>)responseData.get('Payment');

                    if((Integer)paymentData.get('Status') == 0){
                        return;
                    }

                    p = new Payment();
                    p.AccountId = acc.Id;
                    p.Amount = cart.GrandTotalAmount - cart.ValorCashback__c;
                    p.CurrencyIsoCode = cart.CurrencyIsoCode;
                    p.Email = acc.Email__c;
                    p.Date = paymentData.containsKey('CapturedDate') ? Datetime.valueOf((String)paymentData.get('CapturedDate')) :  Datetime.valueOf((String)paymentData.get('ReceivedDate'));
                    p.EffectiveDate = p.Date;
                    p.GatewayDate = p.Date;
                    p.GatewayRefNumber = (String)paymentData.get('PaymentId');
                    p.GatewayResultCode = String.valueOf((Integer)paymentData.get('ReasonCode'));
                    p.GatewayResultCodeDescription = (String)paymentData.get('ReasonMessage');
                    p.GatewayRefDetails = String.valueOf((String)paymentData.get('ReturnMessage')).left(1000);
                    p.ProcessingMode = 'External';
                    p.Status = 'Draft';

                    switch on paymentType {
                        when 'CreditCard' {
                            cpm = new CardPaymentMethod();
                            Map<String,Object> cardData = (Map<String,Object>)paymentData.get('CreditCard');
                            cpm.CardBin = Integer.valueOf(String.valueOf((String)cardData.get('CardNumber')).left(6));
                            cpm.CardCategory = 'CreditCard';
                            cpm.CardHolderName = (String)cardData.get('Holder');
                            cpm.CardLastFour = Integer.valueOf(String.valueOf((String)cardData.get('CardNumber')).right(4));
                            cpm.ExpiryMonth = Integer.valueOf(String.valueOf((String)cardData.get('ExpirationDate')).split('/')[0]);
                            cpm.ExpiryYear = Integer.valueOf(String.valueOf((String)cardData.get('ExpirationDate')).split('/')[1]);
                            cpm.ProcessingMode = 'External';
                            cpm.Status = 'Active';
                            cpm.AccountId = acc.Id;
                            cpm.Email = acc.Email__c;

                            p.Type = 'Sale';
                            p.Status = 'Processed';
                        }
                        when 'Pix' {
                            pa = new PaymentAuthorization();

                            apm = new AlternativePaymentMethod();
                            apm.AccountId = acc.Id;
                            apm.Email = acc.Email__c;
                            apm.PaymentMethodDetails = paymentType;
                            apm.ProcessingMode = 'External';
                            apm.Status = 'Active';
                        
                            p.Type = 'Capture';
                            //p.Status = 'Pending';

                            pa.Status = 'Pending';
                            pa.AccountId = acc.Id;
                            pa.Amount = cart.GrandTotalAmount - cart.ValorCashback__c;
                            pa.CurrencyIsoCode = cart.CurrencyIsoCode;
                            pa.GatewayAuthCode = (String)paymentData.get('ProviderReturnCode');
                            pa.EffectiveDate = p.EffectiveDate;
                            pa.GatewayRefDetails = (String)paymentData.get('QrCodeString');
                            pa.GatewayDate = p.GatewayDate;
                            pa.GatewayRefNumber = p.GatewayRefNumber;
                            pa.GatewayResultCode = p.GatewayResultCode;
                            pa.GatewayResultCodeDescription = p.GatewayResultCodeDescription;
                            pa.ProcessingMode = p.ProcessingMode;
                            pa.Date = p.Date;
                            Map<String,Object> qrCode = (Map<String,Object>)paymentData.get('QrCode');
                            Integer secondsToAdd = (Integer)qrCode.get('Expiration');
                            System.debug(paymentData);
                            pa.ExpirationDate = p.Date.addSeconds(secondsToAdd);

                            // 2. Decodificar a string Base64 em Blob (Binário)
                            Blob pngBlob = EncodingUtil.base64Decode((String)paymentData.get('QrCodeBase64Image'));

                            // 3. Criar a ContentVersion (Arquivo do Salesforce)
                            cv = new ContentVersion();
                            cv.Title = 'PIX_QR_CODE_' + p.GatewayRefNumber;
                            cv.PathOnClient = cv.Title + '.png'; // Define explicitamente a extensão PNG
                            cv.VersionData = pngBlob;
                            cv.IsMajorVersion = true;
                        }
                    }
                }
                when 'CieloGetPayment' {
                    Map<String,Object> paymentData = (Map<String,Object>)responseData.get('Payment');
                    String paymentId = (String)paymentData.get('PaymentId');
                    switch on (String)paymentData.get('Type') {
                        when 'Pix' {
                            switch on (Integer)paymentData.get('Status') {
                                when 0 { // NotFinished - Falhou a geração do QRCode.
                                    message = 'Nenhuma operacao necessaria';
                                }
                                when 1 { // Authorized - Pagamento apto a ser capturado ou definido como pago.
                                    message = 'Nenhuma operacao necessaria';
                                }
                                when 2 { // PaymentConfirmed - Pagou QRCode; Solicitou estorno; Processou estorno parcial.
                                    handlePaid(paymentId);
                                }
                                when 10{  // Voided - Pagamento cancelado.
                                    handleCanceled(paymentId);
                                }
                                when 11 { // Refunded - Processou estorno total. Pagamento cancelado após 23h59 do dia de autorização. Para transações Pix com devolução do valor total, o status sempre será 11
                                    handleRefund(paymentId);
                                }
                                when 12 { // Pending - Gerou QRCode.
                                    message = 'Nenhuma operacao necessaria';
                                }
                                when 13 { // Aborted - Pagamento cancelado por falha no processamento ou por ação do Antifraude.
                                    handleCanceled(paymentId);
                                }
                                when else {
                                    message = 'Nenhuma operacao necessaria';
                                }
                            }
                        }
                        when 'CreditCard' {
                            switch on (Integer)paymentData.get('Status') {
                                when 0 { // NotFinished - Aguardando atualização de status.
                                    message = 'Nenhuma operacao necessaria';
                                }
                                when 1 { // Authorized - Pagamento apto a ser capturado ou definido como pago.
                                    message = 'Nenhuma operacao necessaria';
                                }
                                when 2 { // PaymentConfirmed - Pagamento confirmado e finalizado.
                                    handlePaid(paymentId);
                                }
                                when 3 { // Denied - Pagamento negado por Autorizador.
                                    handleCanceled(paymentId);
                                }
                                when 10 { // Voided - Pagamento cancelado.
                                    handleCanceled(paymentId);
                                }
                                when 11 { // Refunded - Pagamento cancelado após 23h59 do dia de autorização
                                    handleRefund(paymentId);
                                }
                                when 12 { // Pending - 	Aguardando retorno da instituição financeira.
                                    message = 'Nenhuma operacao necessaria';
                                }
                                when 13 { // Aborted - 	Pagamento cancelado por falha no processamento ou por ação do Antifraude.
                                    handleCanceled(paymentId);
                                }
                                when 20 { // Scheduled - Recorrência agendada.
                                    message = 'Nenhuma operacao necessaria';
                                }
                                when else {
                                    message = 'Nenhuma operacao necessaria';
                                }
                            }
                        }
                        when else {
                            
                        }
                    }
                }
                when 'CieloRefund' {
                    if((Integer)responseData.get('Status') == 11 || (String)responseData.get('ReturnCode') == '0'){
                        handleRefund(this.data);
                    }
                }
                when else {
                    System.debug('Invalid Operaton');
                }
            }
        }
        catch (Exception e) {
            LogErro__c l = new LogErro__c(TipoAutomacao__c='Apex', NomeAutomacao__c = 'CommerceCieloIntegration', Status__c = 'Erro', Nome__c = 'Log de Integração Cielo - Processamento da Resposta',TextoErro__c = e.getMessage() + '\n' + e.getStackTraceString());
            logs.add(l);
        }
    }

    private void handlePaid(String paymentId){
        p = [SELECT Id, Status, PaymentGroup.SourceObjectId, PaymentAuthorizationId FROM Payment WHERE GatewayRefNumber = :paymentId LIMIT 1][0];
        if(String.isNotBlank(p.PaymentAuthorizationId)){
            pa = [SELECT Id FROM PaymentAuthorization WHERE Id = :p.PaymentAuthorizationId LIMIT 1][0];
        }

        if(p.Status != 'Processed'){
            message = 'Payment Paid';
            if(pa != null){
                pa.Status = 'Processed';
            }
            p.Status = 'Processed';
            o = new Order();
            o.Id = p.PaymentGroup.SourceObjectId;
            o.Status = PAID_ORDER_STATUS;
        }
        else{
            p = null;
            message = 'Payment already Paid';
        }
    }
    
    private void handleRefund(String paymentId){
        p = [SELECT Id, Status, Amount, Balance, AccountId, PaymentMethodId, PaymentAuthorizationId, PaymentGroupId, OrderPaymentSummaryId, PaymentGroup.SourceObjectId, CurrencyIsoCode FROM Payment WHERE GatewayRefNumber = :paymentId LIMIT 1][0];

        if(p.Balance != 0){
            message = 'Payment Refunded';
            r = new Refund();
            r.Amount = p.Amount;
            r.Type = 'NonReferenced';
            r.CurrencyIsoCode = p.CurrencyIsoCode;
            r.ProcessingMode = 'External';
            r.AccountId = p.AccountId;
            r.PaymentMethodId = p.PaymentMethodId;
            r.EffectiveDate = Datetime.now();
            r.PaymentGroupId = p.PaymentGroupId;
            r.OrderPaymentSummaryId = p.OrderPaymentSummaryId;
            r.Status = 'Processed';

            rlp = new RefundLinePayment();
            rlp.PaymentId = p.Id;
            rlp.HasBeenUnapplied = 'No';
            rlp.Type = 'Applied';
            rlp.Amount = p.Amount;
            rlp.AssociatedAccountId = p.AccountId;
            rlp.CurrencyIsoCode = p.CurrencyIsoCode;

            o = new Order();
            o.Id = p.PaymentGroup.SourceObjectId;
            o.Status = REFUNDED_ORDER_STATUS;
        }
        else{
            p = null;
            message = 'Payment already Refunded';
        }
    }

    private void handleCanceled(String paymentId){
        p = [SELECT Id, Status, PaymentAuthorizationId, PaymentGroup.SourceObjectId, Balance FROM Payment WHERE GatewayRefNumber = :paymentId LIMIT 1][0];
        if(String.isNotBlank(p.PaymentAuthorizationId)){
            pa = [SELECT Id, Amount, CurrencyIsoCode, AccountId FROM PaymentAuthorization WHERE Id = :p.PaymentAuthorizationId LIMIT 1][0];
        }

        if(p.Status != 'Canceled'){
            message = 'Payment Canceled';
            if(pa != null){
                pa.Status = 'Processed';
                /* bug
                paa = new PaymentAuthAdjustment();
                paa.PaymentAuthorizationId = pa.Id;
                paa.Status = 'Pending';
                paa.Type = 'Reversal';
                paa.Amount = pa.Amount;
                paa.CurrencyIsoCode = pa.CurrencyIsoCode;
                paa.AccountId = pa.AccountId;
                paa.EffectiveDate = Datetime.now();
                paa.ProcessingMode = 'External';
                paa.Comments = 'Cancelado Cielo';
                */
                cancelPendingPayment = true;
            }
            p.Status = 'Canceled';

            o = new Order();
            o.Id = p.PaymentGroup.SourceObjectId;
            o.Status = CANCELED_ORDER_STATUS;
        }
        else{
            p = null;
            pa = null;
            message = 'Payment already Canceled';
        }
    }
    
    public void processFinalDMLs(){
        if(!iLogs.isEmpty()) insert iLogs;
        if(!logs.isEmpty()) insert logs;

        if(pg != null){
            upsert pg;
            cart.PaymentGroupId = pg.Id;
        }
        if(cpm != null){
            upsert cpm;
            cart.PaymentMethodId = cpm.Id;
        }
        if(apm != null){
            upsert apm;
            cart.PaymentMethodId = apm.Id;
        }
        if(pa != null){
            if(pa.Id != null){
                update pa;
                if(cancelPendingPayment){
                    pa.Status = 'Canceled';
                    pa.Comments = 'Cancelado Cielo';
                    update pa;
                }
            }
            else{
                pa.PaymentMethodId = cart.PaymentMethodId;
                pa.PaymentGroupId = pg.Id;
                insert pa;
                if(cv != null){
                insert cv;
                
                // 4. Buscar o ContentDocumentId gerado automaticamente
                cv = [SELECT Title, ContentDocumentId FROM ContentVersion WHERE Id = :cv.Id LIMIT 1];

                // 5. Vincular o arquivo ao registro atual (ContentDocumentLink)
                ContentDocumentLink cdl = new ContentDocumentLink();
                cdl.ContentDocumentId = cv.ContentDocumentId;
                cdl.LinkedEntityId = acc.Id;
                cdl.ShareType = 'V'; // Permissão de visualização
                insert cdl;

                // 6. Criar link de distribuição pública para o Experience Site
                ContentDistribution cd = new ContentDistribution();
                cd.Name = cv.Title;
                cd.ContentVersionId = cv.Id;
                cd.PreferencesAllowViewInBrowser = true;
                cd.PreferencesNotifyOnVisit = false;
                insert cd;

                pa.QRCodeUrl__c =[SELECT ContentDownloadUrl FROM ContentDistribution WHERE Id = :cd.Id LIMIT 1].ContentDownloadUrl;
                update pa;
            }
        }
        /* bug
        if(paa != null){
            insert paa;
        }
        */  
        }
        if(p != null){
            if(p.Id != null){
                update p;
            }
            else{
                p.PaymentGroupId = cart.PaymentGroupId;
                p.PaymentMethodId = cart.PaymentMethodId;
                if(pa != null) p.PaymentAuthorizationId = pa.Id;
                insert p;
            }
        }
        if(r != null){
            insert r;
            if(rlp != null){
                rlp.RefundId = r.Id;
                insert rlp;
            }
        }
        if(o != null){
            if(String.isNotBlank(o.Id)){
                update o;
            }
        }
    }

    public void process(){
        createPayload();
        sendPayload();
        processReturn();
        processFinalDMLs();   
    }
    
    /*
    public static String getCardBrand(String bin){
        Integration__mdt param = [SELECT Endpoint__c, Method__c, Timeout__c FROM Integration__mdt WHERE DeveloperName = 'CieloCardBin' LIMIT 1];
        HttpRequest req = new HttpRequest();
        req.setEndpoint(param.Endpoint__c.replace('{BIN}', bin));
        req.setMethod(param.Method__c);
        req.setTimeout(Integer.valueOf(param.Timeout__c));  
        req.setHeader('Content-Type', 'application/json');
        HttpResponse response = new HTTP().send(req);
        Map<String,Object> rData = (Map<String,Object>)JSON.deserializeUntyped(response.getBody());
        return (String)rData.get('Provider');
    }
    */
    /*    
    @AuraEnabled
    public static Map<String,Object> verifyCard(Map<String,Object> data){
        CommerceCieloIntegration integration;
        try {
            integration = new CommerceCieloIntegration('CieloVerifyCard');
            integration.inputData = data;
            integration.createPayload();
            integration.sendPayload();
            integration.processReturn();
            if(!integration.iLogs.isEmpty()) insert integration.iLogs;
            if(!integration.logs.isEmpty()) insert integration.logs;
        } catch (Exception e) {
            throw new AuraHandledException(e.getMessage() + e.getStackTraceString());
        }
        return integration.responseData;
    }
    */
    /*
    @AuraEnabled
    public static Map<String,Object> createPayment(Map<String,Object> data){
        CommerceCieloIntegration integration;
        try {
            integration = new CommerceCieloIntegration('CieloCreatePayment');
            integration.inputData = data;
            integration.createPayload();
            integration.sendPayload();
            integration.processReturn();
            if(!integration.iLogs.isEmpty()) insert integration.iLogs;
            if(!integration.logs.isEmpty()) insert integration.logs;
        } catch (Exception e) {
            throw new AuraHandledException(e.getMessage() + e.getStackTraceString());
        }
        return integration.responseData;
    }
        */

    public class FlowInput {
		@InvocableVariable(required=true label='Nome da Operação' description='CieloVerifyCard,CieloCreatePayment,CieloGetPayment,CieloRefund')
		public String operation;
		@InvocableVariable(label='Dados Adicionais' description='BIN do cartão ou Cielo PaymentId')
		public String data;
        @InvocableVariable(label='Web Cart')
		public WebCart cart;
        @InvocableVariable(label='Account')
		public Account acc;
        @InvocableVariable(label='Cart Delivery Group')
		public CartDeliveryGroup cdg;
        @InvocableVariable(label='Async')
		public Boolean async = false;
        @InvocableVariable(label='Cartão' description='Dados do Cartão de Crédito ou Débito')
		public Card cardInput;
        @InvocableVariable(label='Quantidade de Parcelas')
		public Integer installments;

	}

    @InvocableMethod(label='Cielo Payment Integration')
    public static void invocableExecuteIntegration(List<FlowInput> inputs){
        if(inputs[0].async){
            executeIntegrationAsync(JSON.serialize(inputs));
        }
        else{
            executeIntegration(inputs);
        }
    }

    @future(callout=true)
    public static void executeIntegrationAsync(String inputsStr){
        List<FlowInput> inputs = (List<FlowInput>)JSON.deserializeStrict(inputsStr, List<FlowInput>.class);
        executeIntegration(inputs);
    }

    public static void executeIntegration(List<FlowInput> inputs){
        List<CommerceCieloIntegration> ccis = new List<CommerceCieloIntegration>();
        for(FlowInput i : inputs){
            CommerceCieloIntegration cci = new CommerceCieloIntegration(i.operation, i.data);
            if(i.cart != null)cci.cart = i.cart;
            if(i.acc != null)cci.acc = i.acc;
            if(i.cdg != null) cci.cdg = i.cdg;
            if(i.installments != null) cci.installments = i.installments;
            if(i.cardInput != null) cci.cardInput = i.cardInput;
            
            cci.createPayload();
            cci.sendPayload();
            cci.processReturn();
            ccis.add(cci);
        }
        for(CommerceCieloIntegration cci : ccis){
            cci.processFinalDMLs();
        }
    }

    public final String  CARD_NUMBER_VALUE_REGEX = '("CardNumber"\s*:\s*")[^"]*(")';
    
    public String maskSensitiveData(String raw){
        if(String.isBlank(raw) || !(raw.contains('CardNumber') || raw.contains('SecurityCode'))) return raw;
        String masked = raw;
        Pattern p = Pattern.compile(CARD_NUMBER_VALUE_REGEX);
        Matcher m = p.matcher(masked);
        String rawCardNumberValue = '****************';
        if (m.find()) 
        {
            rawCardNumberValue = m.group(1); 
        }
            masked = masked.replace(CARD_NUMBER_VALUE_REGEX, '"CardNumber":"' + rawCardNumberValue.left(6) + '******' + rawCardNumberValue.right(4) + '"');
            masked = masked.replace('"SecurityCode"\s*:\s*")[^"]*(")', '"SecurityCode":"***"');
        return masked;
    }
}