@IsTest
public with sharing class PagBankPublicKeyTest {
    
    @TestSetup
    static void makeData() {
        
    }

    @IsTest
    public static void CreatePublicKeySuccess01() {
        String publicKey = 'MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnc4jeuW/23oz2/l9WPeO1RLE/yT1/x+WfsDmvibRI8X8rwkW/12swEoD5QqyViphdwwlI6fgldsa0NzjzFFVWGnahKFPIC5UJYTjxMj/h72quxgnJ4KveCBBy+U9WcUo9utaUjQAE4CegvcDLv5WbYDdq2mDNPPNo3kw5MmzhNrPTT0s/veVSTtItMS9fkV29vyFoVVuGCcaLXwE4Fuf1VaYYLjjPgvb+rEoHODUVRMmsBwTuyn5OUMj5bEKwyaw1b82dr2tpxY8lNNbPV5L1CN7mPMZkNcgcty5diq8f2ouRcFX78CW9gRfYRQ3COwhJPQiLzjptRHRsbdwIDAQAB';
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"public_key": "'+ publicKey + '", "created_at": 1580237849044}');
        mock.setStatusCode(201);
        test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        Map<String, Object> mapPublicKey;
        mapPublicKey = (Map<String, Object>) System.JSON.deserializeUntyped(PagBankPublicKey.CreatePublicKey());
        System.assertEquals(publicKey, mapPublicKey.get('publicKey'));
        Test.stopTest();
    }

    @IsTest
    public static void CreatePublicKeyError01() {
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"public_key": istest, "created_at": 1580237849044}'); // JSON Incorrect.
        mock.setStatusCode(201);
        test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        Map<String, Object> mapPublicKey;
        mapPublicKey = (Map<String, Object>) System.JSON.deserializeUntyped(PagBankPublicKey.CreatePublicKey());
        System.assertEquals(true, mapPublicKey.get('hasError'));
        Test.stopTest();
    }

    @IsTest
    public static void getPublicKeySuccess01() {
        String publicKey = 'MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnc4jeuW/23oz2/l9WPeO1RLE/yT1/x+WfsDmvibRI8X8rwkW/12swEoD5QqyViphdwwlI6fgldsa0NzjzFFVWGnahKFPIC5UJYTjxMj/h72quxgnJ4KveCBBy+U9WcUo9utaUjQAE4CegvcDLv5WbYDdq2mDNPPNo3kw5MmzhNrPTT0s/veVSTtItMS9fkV29vyFoVVuGCcaLXwE4Fuf1VaYYLjjPgvb+rEoHODUVRMmsBwTuyn5OUMj5bEKwyaw1b82dr2tpxY8lNNbPV5L1CN7mPMZkNcgcty5diq8f2ouRcFX78CW9gRfYRQ3COwhJPQiLzjptRHRsbdwIDAQAB';
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"public_key": "'+ publicKey + '", "created_at": 1580237849044}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        Map<String, Object> mapPublicKey;
        mapPublicKey = (Map<String, Object>) System.JSON.deserializeUntyped(PagBankPublicKey.getPublicKey());
        System.assertEquals(publicKey, mapPublicKey.get('publicKey'));
        Test.stopTest();
    }

    @IsTest
    public static void getPublicKeyError01() {
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"public_key": istest, "created_at": 1580237849044}'); // JSON Incorrect.
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        Map<String, Object> mapPublicKey;
        mapPublicKey = (Map<String, Object>) System.JSON.deserializeUntyped(PagBankPublicKey.getPublicKey());
        System.assertEquals(true, mapPublicKey.get('hasError'));
        Test.stopTest();
    }

    @IsTest
    public static void changePublicKeySuccess01() {
        String publicKey = 'MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAnc4jeuW/23oz2/l9WPeO1RLE/yT1/x+WfsDmvibRI8X8rwkW/12swEoD5QqyViphdwwlI6fgldsa0NzjzFFVWGnahKFPIC5UJYTjxMj/h72quxgnJ4KveCBBy+U9WcUo9utaUjQAE4CegvcDLv5WbYDdq2mDNPPNo3kw5MmzhNrPTT0s/veVSTtItMS9fkV29vyFoVVuGCcaLXwE4Fuf1VaYYLjjPgvb+rEoHODUVRMmsBwTuyn5OUMj5bEKwyaw1b82dr2tpxY8lNNbPV5L1CN7mPMZkNcgcty5diq8f2ouRcFX78CW9gRfYRQ3COwhJPQiLzjptRHRsbdwIDAQAB';
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"public_key": "'+ publicKey + '", "created_at": 1580237849044}');
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        Map<String, Object> mapPublicKey;
        mapPublicKey = (Map<String, Object>) System.JSON.deserializeUntyped(PagBankPublicKey.changePublicKey());
        System.assertEquals(publicKey, mapPublicKey.get('publicKey'));
        Test.stopTest();
    }

    @IsTest
    public static void changePublicKeyError01() {
        HttpCalloutMockTools mock = new HttpCalloutMockTools();
		mock.setResponseBody('{"public_key": istest, "created_at": 1580237849044}'); // JSON Incorrect.
        mock.setStatusCode(200);
        test.setMock(HttpCalloutMock.class, mock);

        Test.startTest();
        Map<String, Object> mapPublicKey;
        mapPublicKey = (Map<String, Object>) System.JSON.deserializeUntyped(PagBankPublicKey.changePublicKey());
        System.assertEquals(true, mapPublicKey.get('hasError'));
        Test.stopTest();
    }
}