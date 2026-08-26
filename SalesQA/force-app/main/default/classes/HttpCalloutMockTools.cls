@IsTest
public with sharing class HttpCalloutMockTools implements HttpCalloutMock {
    Integer responseStatusCode;
    String responseBody;
    Map<String, String> responseBodyForRequestUrl;
    Map<String, String> responseHeaders;

    public void setStatusCode(Integer statusCode) {
        this.responseStatusCode = statusCode;
    }

    public void setResponseBody(String body) {
        this.responseBody = body;
    }

    public void addResponseBodyForRequest(String url, String body){
        if(this.responseBodyForRequestUrl == null)
            this.responseBodyForRequestUrl = new Map<String, String>();

        this.responseBodyForRequestUrl.put(url, body);
    }

    public void addHeader(String key, String value) {
        if(this.responseHeaders == null)
            this.responseHeaders = new Map<String, String>();

        this.responseHeaders.put(key, value);
    }

    public System.HttpResponse respond(System.HttpRequest request) {
        HttpResponse response = new HttpResponse();
        if(this.responseHeaders != null) {
            for(String key : this.responseHeaders.keySet())
                response.setHeader(key, this.responseHeaders.get(key));
        }

        if(this.responseBodyForRequestUrl != null && this.responseBodyForRequestUrl.containsKey(request.getEndpoint())) {
            String  bodyToSend = this.responseBodyForRequestUrl.get(request.getEndpoint());
            response.setBody(bodyToSend);
        } else 
            response.setBody(this.responseBody);

            response.setStatusCode(this.responseStatusCode);
            return response;
    }
}