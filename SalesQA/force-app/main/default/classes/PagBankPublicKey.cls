public without sharing class PagBankPublicKey {
    private static Map<String, String> mapCalloutMethodXDirectory = new Map<String, String> {
        'POST'  => '/public-keys',
        'GET'   => '/public-keys/card',
        'PUT'   => '/public-keys/card'
    };

    @AuraEnabled
    public static String createPublicKey() {
        String CALLOUTMETHOD = 'POST';
        WrapperReturn newWrapperReturn = new WrapperReturn();
        try {
            CreatePublicKeyRequestWrapper newRequestWrapper = new CreatePublicKeyRequestWrapper();
            newRequestWrapper.type = 'card';

            CalloutWrapper newCalloutWrapper = new CalloutWrapper();
            newCalloutWrapper.json = System.JSON.serialize(newRequestWrapper);
            newCalloutWrapper.calloutMethod = CALLOUTMETHOD;

            HttpResponse response = PagBankPublicKey.executeCallout(newCalloutWrapper);
            ResponseWrapper newResponseWrapper = (ResponseWrapper) System.JSON.deserialize(response.getBody(), PagBankPublicKey.ResponseWrapper.class);
            if(newResponseWrapper != null && String.isNotBlank(newResponseWrapper.public_key)) {
                newWrapperReturn.publicKey = newResponseWrapper.public_key;
            }
        } catch (Exception ex) {
            String errorMessage = 'Class: PagBankPublicKey \n Method: createPublicKey \n Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
            newWrapperReturn.hasError = true;
            newWrapperReturn.errorMessage = errorMessage;
        }
        return System.JSON.serialize(newWrapperReturn);
    }

    @AuraEnabled
    public static String getPublicKey() {
        String CALLOUTMETHOD = 'GET';
        WrapperReturn newWrapperReturn = new WrapperReturn();
        try {
            CalloutWrapper newCalloutWrapper = new CalloutWrapper();
            newCalloutWrapper.calloutMethod = CALLOUTMETHOD;

            HttpResponse response = PagBankPublicKey.executeCallout(newCalloutWrapper);
            ResponseWrapper newResponseWrapper = (ResponseWrapper) System.JSON.deserialize(response.getBody(), PagBankPublicKey.ResponseWrapper.class);
            if(newResponseWrapper != null && String.isNotBlank(newResponseWrapper.public_key)) {
                newWrapperReturn.publicKey = newResponseWrapper.public_key;
            }
        } catch (Exception ex) {
            String errorMessage = 'Class: PagBankPublicKey \n Method: getPublicKey \n Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
            newWrapperReturn.hasError = true;
            newWrapperReturn.errorMessage = errorMessage;
        }
        return System.JSON.serialize(newWrapperReturn);
    }

    @AuraEnabled
    public static String changePublicKey() {
        String CALLOUTMETHOD = 'PUT';
        WrapperReturn newWrapperReturn = new WrapperReturn();
        try {
            CalloutWrapper newCalloutWrapper = new CalloutWrapper();
            newCalloutWrapper.calloutMethod = CALLOUTMETHOD;

            HttpResponse response = PagBankPublicKey.executeCallout(newCalloutWrapper);
            ResponseWrapper newResponseWrapper = (ResponseWrapper) System.JSON.deserialize(response.getBody(), PagBankPublicKey.ResponseWrapper.class);
            if(newResponseWrapper != null && String.isNotBlank(newResponseWrapper.public_key)) {
                newWrapperReturn.publicKey = newResponseWrapper.public_key;
            }
        } catch (Exception ex) {
            String errorMessage = 'Class: PagBankPublicKey \n Method: changePublicKey \n Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
            newWrapperReturn.hasError = true;
            newWrapperReturn.errorMessage = errorMessage;
        }
        return System.JSON.serialize(newWrapperReturn);
    }

    private static HttpResponse executeCallout(PagBankPublicKey.CalloutWrapper calloutWrapper) {
        WebService__mdt service = WebService__mdt.getInstance('PagSeguro');
        Http http = new Http();
        HttpRequest request = new HttpRequest();
        HttpResponse response = new HttpResponse();
        request.setHeader('Content-type', 'application/json');
        // request.setHeader('Authorization', service.Password__c);
        request.setHeader('Authorization', 'Bearer ' + service.AcessToken__c);
        request.setEndpoint(service.BaseUrl__c + mapCalloutMethodXDirectory.get(calloutWrapper.calloutMethod));
        request.setMethod(calloutWrapper.calloutMethod);
        if(String.isNotBlank(calloutWrapper.json)) {
            request.setBody(calloutWrapper.json);
        }
        response = http.send(request);
        createIntegrationLog(request.getEndpoint(), request.getMethod(), response.getStatusCode(), response.getStatus(), request.getBody(), response.getBody());
        return response;
    }

    @Future
    private static void createIntegrationLog(String endPoint, String integrationMethod, Integer integrationStatusCode, String integrationStatus, String requestBody, String responseBody) {
        IntegrationLog__c log = new IntegrationLog__c(
            NomeLog__c = 'PUBLIC KEY',
			ApexClass__c = 'PagBankPublicKey',
            ApexMethod__c = 'createIntegrationLog',
			EndPoint__c = endPoint,
			IntegrationMethod__c = integrationMethod,
			IntegrationStatusCode__c = integrationStatusCode + ' : ' + integrationStatus,
			RequestBody01__c = requestBody,
			OperationStatus__c = integrationStatusCode != 200 && integrationStatusCode != 201 ? 'Error' : 'Success',
            responseBody01__c =  responseBody
        );
        insert log;
    }

    private class CreatePublicKeyRequestWrapper {
        String type;    
    }
    
    private class ResponseWrapper {
        String public_key;
        Blob created_at;
    }
    
    private class CalloutWrapper {
        String json;
        String calloutMethod;   
    }

    private class WrapperReturn {
        Boolean hasError = false;
        String errorMessage;
        String publicKey;
    }
}