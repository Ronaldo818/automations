@IsTest
public with sharing class BuyerAccountHandlerTest {

    @IsTest
    static void validatePendingCommerceRecordsSuccess01() {
        Test.startTest();
        Account newAccount = new Account(
            Name = 'isTest',
            Unidade__c = 1,
            Empresa__c = 1,
            CodigoClienteERP__c = '1',
            Ocorrencia__c = 1,
            Lista__c = 27,
            ClienteEcommerce__c = true
        );
        insert newAccount;
        Test.stopTest();

        BuyerAccount buyerAcc = [
            SELECT Id, PendingEcommerceRecords__c
            FROM BuyerAccount
            WHERE BuyerId = :newAccount.Id
        ];

        System.assertEquals(true, buyerAcc.PendingEcommerceRecords__c);
    }
}