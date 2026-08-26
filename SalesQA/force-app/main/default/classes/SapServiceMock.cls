@isTest
global class SapServiceMock implements HttpCalloutMock {
    public SapServiceMock() {

    }

    global HTTPResponse respond(HTTPRequest req) {
            
        if(req.getEndpoint().contains('/salesorder')){
            // Create Transaction Confirmation Response
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');
            res.setBody('{"A_SalesOrder":{"SalesOrder":"179","SalesOrderType":"ORB","CreationDate":"2025-07-24T00:00:00.000","CreatedByUser":"S-CPI","CustomerPaymentTerms":"0009","to_Item":[{"SalesOrderItemText":"Prod.comerc.10,PD,marca registrada","MaterialGroup":"M010","CustomerPaymentTerms":"0009","PurchaseOrderByCustomer":"00001","ItemNetWeight":"405.000","NetAmount":"2508.21","SalesOrderItemCategory":"TBN","Material":"TG11","SalesOrderItem":"10","RequestedQuantity":"25","RequestedQuantityUnit":"UN"}]}}');
            res.setStatusCode(200);
            return res;  
        }    
        return null;
    }
    
}