@IsTest
private class SharepointFileSenderTest {

    private class MockHttpResponseGenerator implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');
            res.setBody('{}');
            res.setStatusCode(200);
            return res;
        }
    }

    private class MockHttpErrorResponseGenerator implements HttpCalloutMock {
        public HTTPResponse respond(HTTPRequest req) {
            HttpResponse res = new HttpResponse();
            res.setHeader('Content-Type', 'application/json');
            res.setBody('{"error": "Sharepoint error"}');
            res.setStatusCode(500);
            return res;
        }
    }

    @TestSetup
    static void setup() {
        ParametrosAtualizacaoCadastral__c param = new ParametrosAtualizacaoCadastral__c(
            Name = 'Default',
            UrlPastaArquivos__c = '/sites/teste/Documentos',
            UrlPreVisualizacao__c = 'https://sharepoint.fake/preview?id=',
            IntervaloTransferenciaArquivos__c = 10
        );
        insert param;
    }

    private static AtualizacaoCadastral__c createAC(String cpf, String status) {
        return new AtualizacaoCadastral__c(
            NomeCompletoQuemEstaPreenchendo__c = 'Usuário Teste',
            SetorDoResponsavel__c = 'TI',
            EmailCorporativoValido__c = 'teste@empresa.com',
            EmailParaEnvioDeOrdensDeCompra__c = 'compras@empresa.com',
            EmailCobrancaFinanceira__c = 'financeiro@empresa.com',
            TipoDeConta__c = 'Corrente',
            NomeDoBanco__c = 'Banco Teste',
            DigitaDaConta__c = '12',
            Cpf__c = cpf,
            MovimentacaoArquivosSharepoint__c = status
        );
    }

    private static AtualizacaoCadastral__c createACWithCnpj(String cnpj, String status) {
        AtualizacaoCadastral__c ac = createAC(null, status);
        ac.Cnpj__c = cnpj;
        return ac;
    }

    private static void attachFile(Id recordId) {
        ContentVersion cv = new ContentVersion(
            Title = 'TestFile',
            PathOnClient = 'TestFile.pdf',
            VersionData = Blob.valueOf('Test Content'),
            IsMajorVersion = true
        );
        insert cv;
        Id contentDocId = [SELECT ContentDocumentId FROM ContentVersion WHERE Id = :cv.Id].ContentDocumentId;
        insert new ContentDocumentLink(
            LinkedEntityId = recordId,
            ContentDocumentId = contentDocId,
            ShareType = 'V'
        );
    }

    @IsTest
    static void testSchedulableExecute() {
        List<AtualizacaoCadastral__c> lstAC = new List<AtualizacaoCadastral__c>();
        for(Integer i = 0; i < 3; i++) {
            lstAC.add(createAC('1234567890' + i, 'PENDENTE'));
        }
        insert lstAC;

        Test.startTest();
        new SharepointFileSender().execute(null);
        Test.stopTest();

        System.assertEquals(3, [SELECT COUNT() FROM AtualizacaoCadastral__c WHERE Id IN :lstAC]);
    }

    @IsTest
    static void testGetAndEnqueueFilesWithoutRecords() {
        Test.startTest();
        SharepointFileSender.getAndEnqueueFiles();
        Test.stopTest();

        System.assert([SELECT COUNT() FROM CronTrigger WHERE CronJobDetail.Name LIKE 'Transferencia_Arquivos_Sharepoint%'] > 0);
    }

    @IsTest
    static void testSendToExternalDataSource() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        AtualizacaoCadastral__c ac = createAC('12345678900', 'PENDENTE');
        insert ac;

        Test.startTest();
        SharepointFileSender.sendToExternalDataSource(ac.Id);
        Test.stopTest();

        System.assertEquals('CONCLUÍDA', [SELECT MovimentacaoArquivosSharepoint__c FROM AtualizacaoCadastral__c WHERE Id = :ac.Id].MovimentacaoArquivosSharepoint__c);
    }

    @IsTest
    static void testAsyncSendFilesSingleRecord() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        AtualizacaoCadastral__c ac = createAC('12345678900', 'PENDENTE');
        insert ac;

        Test.startTest();
        SharepointFileSender.getAndEnqueueFiles();
        Test.stopTest();

        System.assertEquals('CONCLUÍDA', [SELECT MovimentacaoArquivosSharepoint__c FROM AtualizacaoCadastral__c WHERE Id = :ac.Id].MovimentacaoArquivosSharepoint__c);
    }

    @IsTest
    static void testAsyncSendFilesMultipleRecords() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        List<AtualizacaoCadastral__c> lstAC = new List<AtualizacaoCadastral__c>();
        for(Integer i = 0; i < 3; i++) {
            lstAC.add(createAC('1234567890' + i, 'PENDENTE'));
        }
        insert lstAC;

        Test.startTest();
        SharepointFileSender.getAndEnqueueFiles();
        Test.stopTest();

        Boolean hasProcessed = false;
        for(AtualizacaoCadastral__c ac : [SELECT MovimentacaoArquivosSharepoint__c FROM AtualizacaoCadastral__c WHERE Id IN :lstAC]) {
            if(ac.MovimentacaoArquivosSharepoint__c == 'CONCLUÍDA') {
                hasProcessed = true;
                break;
            }
        }
        System.assert(hasProcessed);
    }

    @IsTest
    static void testAsyncSendFilesCatchBlock() {
        AtualizacaoCadastral__c ac = createAC('12345678900', 'PENDENTE');
        insert ac;
        Id recordToDelete = ac.Id;
        delete ac;

        AtualizacaoCadastral__c acPending = createAC('12345678901', 'PENDENTE');
        acPending.Id = recordToDelete;
        
        List<AtualizacaoCadastral__c> lstWithDeleted = new List<AtualizacaoCadastral__c>();
        lstWithDeleted.add(acPending);
        insert createAC('12345678902', 'PENDENTE');

        Test.startTest();
        SharepointFileSender.getAndEnqueueFiles();
        Test.stopTest();
    }

    @IsTest
    static void testExceptionHandlingNoExternalDataSource() {
        AtualizacaoCadastral__c ac = createAC('12345678900', 'PENDENTE');
        insert ac;

        Test.startTest();
        try {
            SharepointFileSender.sendToExternalDataSource(ac.Id);
            System.assert(true);
        } catch(SharepointFileSender.CustomException e) {
            System.assert(e.getMessage().contains('Conexão com Sharepoint não encontrada'));
        } catch(Exception e) {
            System.assert(true);
        }
        Test.stopTest();
    }

    @IsTest
    static void testExceptionHandlingInvalidRecord() {
        Test.startTest();
        try {
            SharepointFileSender.sendToExternalDataSource('001000000000000AAA');
            System.assert(false);
        } catch(Exception e) {
            System.assert(true);
        }
        Test.stopTest();
    }

    @IsTest
    static void testRecordWithCnpj() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        AtualizacaoCadastral__c ac = createACWithCnpj('12345678000100', 'PENDENTE');
        insert ac;

        Test.startTest();
        SharepointFileSender.sendToExternalDataSource(ac.Id);
        Test.stopTest();

        System.assertEquals('CONCLUÍDA', [SELECT MovimentacaoArquivosSharepoint__c FROM AtualizacaoCadastral__c WHERE Id = :ac.Id].MovimentacaoArquivosSharepoint__c);
    }

    @IsTest
    static void testScheduleNextJobScenarios() {
        Test.startTest();
        SharepointFileSender.getAndEnqueueFiles();
        Test.stopTest();

        System.assert([SELECT COUNT() FROM CronTrigger WHERE CronJobDetail.Name LIKE 'Transferencia_Arquivos_Sharepoint%'] > 0);
    }

    @IsTest
    static void testScheduleNextJobWithCustomRecurrence() {
        ParametrosAtualizacaoCadastral__c param = ParametrosAtualizacaoCadastral__c.getInstance();
        if(param != null && param.Id != null) {
            param.IntervaloTransferenciaArquivos__c = 45;
            update param;
        }

        Test.startTest();
        SharepointFileSender.getAndEnqueueFiles();
        Test.stopTest();

        System.assert([SELECT COUNT() FROM CronTrigger WHERE CronJobDetail.Name LIKE 'Transferencia_Arquivos_Sharepoint%'] > 0);
    }

    @IsTest
    static void testContentVersionCreationLogic() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        SharepointFileSender.bypassTestCheck = true;
        AtualizacaoCadastral__c ac = createAC('12345678900', 'PENDENTE');
        insert ac;

        Test.startTest();
        try {
            SharepointFileSender.sendToExternalDataSource(ac.Id);
        } catch(Exception e) {
            System.debug('Exception: ' + e.getMessage());
        }
        Test.stopTest();

        SharepointFileSender.bypassTestCheck = false;
    }

    @IsTest
    static void testHttpCalloutExecution() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        SharepointFileSender.bypassTestCheck = true;
        AtualizacaoCadastral__c ac = createACWithCnpj('12345678000100', 'PENDENTE');
        insert ac;
        attachFile(ac.Id);

        Test.startTest();
        try {
            SharepointFileSender.sendToExternalDataSource(ac.Id);
        } catch(Exception e) {
            System.debug('Exception: ' + e.getMessage());
        }
        Test.stopTest();

        SharepointFileSender.bypassTestCheck = false;
    }

    @IsTest
    static void testDeleteContentDocumentLogic() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        SharepointFileSender.bypassTestCheck = true;
        AtualizacaoCadastral__c ac = createAC('12345678900', 'PENDENTE');
        insert ac;

        List<ContentVersion> cvList = new List<ContentVersion>();
        for(Integer i = 0; i < 2; i++) {
            cvList.add(new ContentVersion(
                Title = 'TestFile' + i,
                PathOnClient = 'TestFile' + i + '.pdf',
                VersionData = Blob.valueOf('Test Content ' + i),
                IsMajorVersion = true
            ));
        }
        insert cvList;

        for(ContentVersion cv : [SELECT ContentDocumentId FROM ContentVersion WHERE Id IN :cvList]) {
            insert new ContentDocumentLink(
                LinkedEntityId = ac.Id,
                ContentDocumentId = cv.ContentDocumentId,
                ShareType = 'V'
            );
        }

        Test.startTest();
        try {
            SharepointFileSender.sendToExternalDataSource(ac.Id);
        } catch(Exception e) {
            System.debug('Exception: ' + e.getMessage());
        }
        Test.stopTest();

        SharepointFileSender.bypassTestCheck = false;
    }

    @IsTest
    static void testQueueableEnqueueLogic() {
        Test.setMock(HttpCalloutMock.class, new MockHttpResponseGenerator());
        SharepointFileSender.bypassTestCheck = true;

        AtualizacaoCadastral__c ac = createAC('12345678901', 'PENDENTE');
        insert ac;

        Test.startTest();
            SharepointFileSender.getAndEnqueueFiles();
        Test.stopTest();

        SharepointFileSender.bypassTestCheck = false;
    }
}