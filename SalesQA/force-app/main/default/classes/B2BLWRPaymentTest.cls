@IsTest
public class B2BLWRPaymentTest {
	
    @TestSetup
    static void makeData() {
        
        Payment newPayment = new Payment(
        	Amount 			= 1,
            Status 			= 'Draft',
			Type 			= 'Capture',
            ProcessingMode	= 'External'
        );
		insert newPayment;
        
        PaymentAuthorization newPaymentAuthorization = new PaymentAuthorization(
			Amount 			= 1,
			Status 			= 'Draft',
			ProcessingMode	= 'External'
		);
		insert newPaymentAuthorization;
        
    }
    
    @IsTest
    static void createAuthResponseSuccess() {
        CommercePayments.CaptureRequest captureRequest = new CommercePayments.CaptureRequest(1, [SELECT Id FROM PaymentAuthorization LIMIT 1].Id);
        commercepayments.paymentGatewayContext gatewayContext = new commercepayments.paymentGatewayContext(captureRequest, commercepayments.RequestType.Capture);
        B2BLWRPayment b2bPayment = new B2BLWRPayment();
        
        Test.startTest();
		b2bPayment.processRequest(gatewayContext);
        Test.stopTest();
    }
    
    @IsTest
    static void createCaptureResponseSuccess() {
        CommercePayments.AuthorizationRequest authorizationRequest = new CommercePayments.AuthorizationRequest(1);
        commercepayments.paymentGatewayContext gatewayContext = new commercepayments.paymentGatewayContext(authorizationRequest, commercepayments.RequestType.Authorize);
        B2BLWRPayment b2bPayment = new B2BLWRPayment();
        
        Test.startTest();
		b2bPayment.processRequest(gatewayContext);
        Test.stopTest();   
    }
    
    @IsTest
    static void createSaleResponseSuccess() {
        CommercePayments.SaleRequest SaleRequest = new CommercePayments.SaleRequest(1);
        commercepayments.paymentGatewayContext gatewayContext = new commercepayments.paymentGatewayContext(SaleRequest, commercepayments.RequestType.Sale);
        B2BLWRPayment b2bPayment = new B2BLWRPayment();
        
        Test.startTest();
		b2bPayment.processRequest(gatewayContext);
        Test.stopTest();
    }
    
    @IsTest
    static void refundSuccess() {
        CommercePayments.ReferencedRefundRequest refundRequest = new CommercePayments.ReferencedRefundRequest(1, [SELECT Id FROM Payment LIMIT 1].Id);
        commercepayments.paymentGatewayContext gatewayContext = new commercepayments.paymentGatewayContext(refundRequest, commercepayments.RequestType.ReferencedRefund);
        B2BLWRPayment b2bPayment = new B2BLWRPayment();
        
        Test.startTest();
		b2bPayment.processRequest(gatewayContext);
        Test.stopTest();   
    }
    
    @IsTest
    static void tokenizationSuccess() {
        CommercePayments.PaymentMethodTokenizationRequest tokenizationRequest = new CommercePayments.PaymentMethodTokenizationRequest();
        commercepayments.paymentGatewayContext gatewayContext = new commercepayments.paymentGatewayContext(tokenizationRequest, commercepayments.RequestType.Tokenize);
        B2BLWRPayment b2bPayment = new B2BLWRPayment();
        
        Test.startTest();
		b2bPayment.processRequest(gatewayContext);
        Test.stopTest();   
    }
    
}