@isTest
global class CieloIntegrationMock  implements HttpCalloutMock {
    String errorType;

    String cartId;

    global CieloIntegrationMock(){

    }

    global CieloIntegrationMock(String errorCode){
        errorType = errorCode;
    }

    global void setCartId(String value){
        cartId = value;
    }

    global HTTPResponse respond(HTTPRequest req) {
        HttpResponse res = new HttpResponse();
        res.setHeader('Content-Type', 'application/json');
        if(req.getEndpoint().contains('/1/zeroauth/')){
            switch on errorType {
                when '39'{
                    // Success Return
                    res.setBody('{"Valid":false,"ReturnCode":"39","ReturnMessage":"Autorizacao negada","IssuerTransactionId":"512677902713598"}');
                    res.setStatusCode(200);
                    return res;
                }
                when else {
                    // Success Return
                    res.setBody('{"Valid":true,"ReturnCode":"00","ReturnMessage":"Transacao autorizada","IssuerTransactionId":"847529"}');
                    res.setStatusCode(200);
                    return res;
                }
            }     
        }
        else if(req.getEndpoint().contains('/void')){
            res.setBody('{"Status":2,"ReasonCode":8,"ReasonMessage":"Scheduled","ProviderReturnCode":"0","ProviderReturnMessage":"Devolução solicitada com sucesso","ReturnCode":"0","ReturnMessage":"Devolução solicitada com sucesso","AuthorizationCode":"E01027058202412021928yOpJfcc9zLl","Links":[{"Method":"GET","Rel":"self","Href":"https://apiquerysandbox.cieloecommerce.cielo.com.br/1/sales/b8c1b2ea-e06a-4135-9389-8bdbdccacd20"},{"Method":"PUT","Rel":"void","Href":"https://apisandbox.cieloecommerce.cielo.com.br/1/sales/b8c1b2ea-e06a-4135-9389-8bdbdccacd20/void"}]}');
            res.setStatusCode(200);
            return res;
        }
        else if(req.getEndpoint().contains('/1/sales')){
            switch on errorType {
                when 'pix_success'{
                    // Success Return
                    //res.setBody('{"MerchantOrderId":"E6Z2K-PMG6D-EB6NK-PTP5L","Customer":{"Name":"LARISSA ALVES GOMES","Identity":"13456023677","IdentityType":"CPF","Email":"laryssa811@hotmail.com.invalid","Address":{"Street":"RUA MARIA PINTO DA SILVA, 470","ZipCode":"35502-084","City":"DIVINOPOLIS","Country":"BR","AddressType":0},"DeliveryAddress":{"Street":"RUA MARIA PINTO DA SILVA, 470","ZipCode":"35502-084","City":"DIVINOPOLIS","Country":"BR","AddressType":0}},"Payment":{"QrCodeBase64Image":"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","QrCodeExpiration":3600,"QrCodeString":"00020126370014BR.GOV.BCB.PIX0115pix@example.com5204000053039865406876.545802BR5913FULANO DE TAL6008BRASILIA6229052500000000-0000-0000-0000-063042AB7","Tid":"0622062023479","SentOrderId":"E6Z2K-PMG6D-EB6NK-PTP5L","Amount":11029,"ReceivedDate":"2026-06-22 06:20:23","Provider":"Simulado","ExtraDataCollection":[{"Name":"CartId","Value":"0a6be000000FzrBAAS"}],"Status":12,"IsSplitted":false,"ReturnMessage":"QR Code PIX gerado com sucesso.","ReturnCode":"0","PaymentId":"Test","Type":"Pix","Currency":"BRL","Country":"BRA","Links":[{"Method":"GET","Rel":"self","Href":"https://apiquerysandbox.cieloecommerce.cielo.com.br/1/sales/1698a447-2e2c-4d75-9bb6-5b72ffb37991"}]}}');
                    res.setBody('{"MerchantOrderId":"********************************","Customer":{"Name":"Aline de Souza","Identity":"066....","IdentityType":"CPF"},"Payment":{"QrCode":{"Expiration":86400},"QrCodeBase64Image":"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","QrCodeString":"00020[...]","SentOrderId":"********************************","Amount":100,"ReceivedDate":"2024-11-28 16:29:52","Provider":"Cielo2","Status":12,"IsSplitted":false,"ReturnMessage":"QRCode gerado com sucesso","ReturnCode":"0","PaymentId":"Test","Type":"Pix","Currency":"BRL","Country":"BRA","Links":[{"Method":"GET","Rel":"self","Href":"https://apiquerysandbox.cieloecommerce.cielo.com.br/1/sales/********-****-****-****-************"}]}}');
                    res.setStatusCode(201);
                    return res;
                }
                when else {
                
                    // Success Return
                    //res.setBody('{"MerchantOrderId":"6QLUD-VBZ6J-273KW-WGMB5","Customer":{"Name":"LARISSA ALVES GOMES","Identity":"13456023677","IdentityType":"CPF","Email":"laryssa811@hotmail.com.invalid","Address":{"Street":"RUA MARIA PINTO DA SILVA, 470","ZipCode":"35502-084","City":"DIVINOPOLIS","Country":"BR","AddressType":0},"DeliveryAddress":{"Street":"RUA MARIA PINTO DA SILVA, 470","ZipCode":"35502-084","City":"DIVINOPOLIS","Country":"BR","AddressType":0}},"Payment":{"ServiceTaxAmount":0,"Installments":1,"Interest":0,"Capture":true,"Authenticate":false,"Recurrent":false,"CreditCard":{"CardNumber":"550209******0000","Holder":"TESTE","ExpirationDate":"01/2027","SaveCard":false,"Alias":"","Brand":"Undefined","PaymentAccountReference":"PN9K08PN9491PVLQ0YC909NMOAB7N"},"Tid":"0622064046228","ProofOfSale":"289393","AuthorizationCode":"700220","SoftDescriptor":"MercadoAvivar","Provider":"Simulado","IsQrCode":false,"Tip":false,"SentOrderId":"202606220640467230F1","Amount":11029,"ReceivedDate":"2026-06-22 06:40:46","CapturedAmount":11029,"CapturedDate":"2026-06-22 06:40:46","ExtraDataCollection":[{"Name":"CartId","Value":"0a6be000000FzpZAAS"}],"Status":2,"IsSplitted":false,"ReturnMessage":"Operation Successful","ReturnCode":"6","PaymentId":"Test","Type":"CreditCard","Currency":"BRL","Country":"BRA","Links":[{"Method":"GET","Rel":"self","Href":"https://apiquerysandbox.cieloecommerce.cielo.com.br/1/sales/d3f8cb18-c004-4e43-9f56-da04f34ca934"},{"Method":"PUT","Rel":"void","Href":"https://apisandbox.cieloecommerce.cielo.com.br/1/sales/d3f8cb18-c004-4e43-9f56-da04f34ca934/void"}]}}');
                    res.setBody('{"MerchantOrderId":"2017051001","Customer":{"Name":"Aline de Souza","Identity":"12345678909","IdentityType":"CPF","Email":"aline@email.com","Birthdate":"1990-01-01","Address":{"Street":"Alameda Xingu","Number":"512","Complement":"27 andar","ZipCode":"12345987","City":"São Paulo","State":"SP","Country":"BRA","AddressType":0},"DeliveryAddress":{"Street":"Alameda Xingu","Complement":"27 andar","ZipCode":"12345987","City":"São Paulo","State":"SP","Country":"BRA","AddressType":0}},"Payment":{"ServiceTaxAmount":0,"Installments":1,"Interest":0,"Capture":true,"Authenticate":false,"Recurrent":false,"IssuerTransactionId":"580027442382078","CreditCard":{"CardNumber":"537154******0634","Holder":"Aline de Souza","ExpirationDate":"07/2027","SaveCard":false,"Brand":"Master","CardOnFile":{"Usage":"Used","Reason":"Unscheduled"},"PaymentAccountReference":"ZEYO5L521MBLKDQU0HHXENP1B1B58"},"Tid":"1209032335683","ProofOfSale":"255231","AuthorizationCode":"875176","SoftDescriptor":"LojaTeste","Provider":"Simulado","IsQrCode":false,"InitiatedTransactionIndicator":{"Category":"C1","Subcategory":"StandingOrder"},"Tip":false,"SentOrderId":"202512091523357AF3AB","Amount":10000,"ReceivedDate":"2025-12-09 15:23:35","CapturedAmount":10000,"CapturedDate":"2025-12-09 15:23:35","Status":2,"IsSplitted":false,"ReturnMessage":"Operation Successful","ReturnCode":"6","PaymentId":"Test","Type":"CreditCard","Currency":"BRL","Country":"BRA","Links":[{"Method":"GET","Rel":"self","Href":"https://apiquerysandbox.cieloecommerce.cielo.com.br/1/sales/71c908df-c72d-425c-89d2-3376f1b5a76d"},{"Method":"PUT","Rel":"void","Href":"https://apisandbox.cieloecommerce.cielo.com.br/1/sales/71c908df-c72d-425c-89d2-3376f1b5a76d/void"}]}}');
                    res.setStatusCode(201);
                    return res;
                
                }
            }
        }
        else if(req.getEndpoint().contains('/payments/') && req.getMethod() == 'GET' ){
            // Create Order Response
            
            switch on errorType {
                when 'pix'{
                    // Create Successful Preference Response
                    res.setBody('{"MerchantOrderId":"Loja123456","AcquirerOrderId":"ECO77777MEVKS5WE","Customer":{"Name":"Nome do Pagador","Identity":"12345678909","Address":{}},"Payment":{"Installments":1,"Tid":"c9a447a4-1511-3aea-8a67-6057905e4302","ProofOfSale":"00000075","PaymentId":"Test","Type":"Pix","Amount":100,"ReceivedDate":"2024-12-13 11:14:11","Currency":"BRL","Country":"BRA","Provider":"Cielo","Status":12,"Links":[{"Method":"GET","Rel":"self","Href":"https://apiquerysandbox.cieloecommerce.cielo.com.br/1/sales/47a16937-7275-4088-a7de-e2a9d70f1ca4"}]}}');
                    res.setStatusCode(200);
                    return res;
                }
                when else {
                    // Create Successful Preference Response
                    res.setBody('{"MerchantOrderId":"Loja123456","Customer":{"Name":"Comprador crédito completo","Email":"compradorteste@teste.com","Birthdate":"1991-01-02","Address":{"Street":"Rua Teste","Number":"123","Complement":"AP 123","ZipCode":"12345987","City":"Rio de Janeiro","State":"RJ","Country":"BRA"}},"Payment":{"ServiceTaxAmount":0,"Installments":1,"Interest":"ByMerchant","Capture":true,"Authenticate":false,"Recurrent":false,"CreditCard":{"CardNumber":"123412******1234","Holder":"Teste Holder","ExpirationDate":"12/2030","Brand":"Visa","PaymentAccountReference":"NUKBZZQ3E34FLV2SAHZXJBVSGWWV8"},"ProofOfSale":"553486","Tid":"1212092931654","AuthorizationCode":"520520","PaymentId":"Test","Type":"CreditCard","Amount":15700,"ReceivedDate":"2024-12-12 09:29:31","CapturedAmount":15700,"CapturedDate":"2024-12-12 09:29:31","Currency":"BRL","Country":"BRA","Provider":"Simulado","Status":2,"Links":[{"Method":"GET","Rel":"self","Href":"https://apiquerysandbox.cieloecommerce.cielo.com.br/1/sales/2352fc91-f9a4-4ca2-aedb-31488b9658c9"},{"Method":"PUT","Rel":"void","Href":"https://apisandbox.cieloecommerce.cielo.com.br/1/sales/2352fc91-f9a4-4ca2-aedb-31488b9658c9/void"}]}}');
                    res.setStatusCode(200);
                    return res;
                }
            } 
        }
        return null;
    }
}