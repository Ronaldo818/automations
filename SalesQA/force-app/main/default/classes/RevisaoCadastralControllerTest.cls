@IsTest
private class RevisaoCadastralControllerTest {

    private static final String CNPJ_3_LIMPO = '44555666000181';
    private static final String CNPJ_4_LIMPO = '55666777000181';
    private static final String CNPJ_ALFA_LIMPO = 'AB12CD34000181';
    private static final String CNPJ_ALFA_MASCARADO = 'AB.12C.D34/0001-81';

    // ─── Helpers ────────────────────────────────────────────────────────────

    private static Account createAccount(String cnpj) {
        Account acc = new Account(
            Name = 'Conta Teste',
            CNPJ__c = cnpj,
            CNPJInput__c = cnpj
        );
        insert acc;
        return acc;
    }

    // ─── saveRevisaoCadastral ────────────────────────────────────────────────

    @IsTest
    static void testSaveRevisaoCadastral_CnpjAlfanumerico() {
        Map<String, Object> fields = new Map<String, Object>{
            'Cnpj__c' => CNPJ_ALFA_MASCARADO
        };

        Test.startTest();
        Id revisaoId = RevisaoCadastralController.saveRevisaoCadastral(JSON.serialize(fields));
        Test.stopTest();

        RevisaoCadastral__c revisao = [
            SELECT Cnpj__c
            FROM RevisaoCadastral__c
            WHERE Id = :revisaoId
            LIMIT 1
        ];

        System.assertEquals(CNPJ_ALFA_LIMPO, revisao.Cnpj__c);
    }

    @IsTest
    static void testSaveRevisaoCadastral_CnpjNulo() {
        Map<String, Object> fields = new Map<String, Object>{ 'Cnpj__c' => null };

        Test.startTest();
        Id revisaoId = RevisaoCadastralController.saveRevisaoCadastral(JSON.serialize(fields));
        Test.stopTest();

        RevisaoCadastral__c revisao = [
            SELECT Cnpj__c FROM RevisaoCadastral__c WHERE Id = :revisaoId LIMIT 1
        ];
        System.assertEquals(null, revisao.Cnpj__c, 'CNPJ deve ser nulo quando o valor fornecido é nulo.');
    }

    @IsTest
    static void testSaveRevisaoCadastral_JsonVazio() {
        Test.startTest();
        Id revisaoId = RevisaoCadastralController.saveRevisaoCadastral('{}');
        Test.stopTest();

        System.assertNotEquals(null, revisaoId, 'Deve inserir registro mesmo com JSON vazio.');
    }

    // ─── findAccountByCnpj ──────────────────────────────────────────────────

    @IsTest
    static void testFindAccountByCnpj_Alfanumerico() {
        Account acc = createAccount(CNPJ_ALFA_LIMPO);

        Test.startTest();
        Account result = RevisaoCadastralController.findAccountByCnpj(CNPJ_ALFA_LIMPO);
        Test.stopTest(); 

        System.assertNotEquals(null, result);
        System.assertEquals(acc.Id, result.Id);
    }

    @IsTest
    static void testFindAccountByCnpj_NaoEncontrada() {
        Test.startTest();
        Account result = RevisaoCadastralController.findAccountByCnpj('00.000.000/0000-00');
        Test.stopTest();

        System.assertEquals(null, result, 'Deve retornar null quando não há conta com esse CNPJ.');
    }

    // ─── getTipoRevisaoOptions ───────────────────────────────────────────────

    @IsTest
    static void testGetTipoRevisaoOptions() {
        Test.startTest();
        List<Map<String, String>> options = RevisaoCadastralController.getTipoRevisaoOptions();
        Test.stopTest();

        System.assertNotEquals(null, options, 'A lista de opções não deve ser nula.');

        for (Map<String, String> opt : options) {
            System.assert(opt.containsKey('label'), 'Cada opção deve conter a chave "label".');
            System.assert(opt.containsKey('value'), 'Cada opção deve conter a chave "value".');
            System.assert(String.isNotBlank(opt.get('label')), 'O label não deve ser vazio.');
            System.assert(String.isNotBlank(opt.get('value')), 'O value não deve ser vazio.');
        }
    }

    // ─── saveFiles ───────────────────────────────────────────────────────────

    @IsTest
    static void testSaveFiles_Sucesso() {
        Account acc       = createAccount(CNPJ_3_LIMPO);
        String fileName   = 'documento_teste.pdf';
        String base64Data = EncodingUtil.base64Encode(Blob.valueOf('conteudo do arquivo de teste'));

        Test.startTest();
        RevisaoCadastralController.saveFiles(acc.Id, fileName, base64Data);
        Test.stopTest();

        List<ContentVersion> versions = [
            SELECT Id, Title, PathOnClient
            FROM   ContentVersion
            WHERE  Title = :fileName
        ];
        System.assertEquals(1, versions.size(), 'Deve existir exatamente um ContentVersion.');
        System.assertEquals(fileName, versions[0].PathOnClient, 'PathOnClient incorreto.');

        Id contentDocumentId = [
            SELECT ContentDocumentId FROM ContentVersion WHERE Id = :versions[0].Id LIMIT 1
        ].ContentDocumentId;

        List<ContentDocumentLink> links = [
            SELECT Id, LinkedEntityId, ShareType
            FROM   ContentDocumentLink
            WHERE  ContentDocumentId = :contentDocumentId
              AND  LinkedEntityId    = :acc.Id
        ];
        System.assertEquals(1, links.size(), 'Deve existir exatamente um ContentDocumentLink.');
        System.assertEquals('V', links[0].ShareType, 'ShareType deve ser "V".');
    }

    @IsTest
    static void testSaveFiles_MultiplosArquivos() {
        Account acc       = createAccount(CNPJ_4_LIMPO);
        String base64Data = EncodingUtil.base64Encode(Blob.valueOf('dados'));

        Test.startTest();
        RevisaoCadastralController.saveFiles(acc.Id, 'arquivo1.txt', base64Data);
        RevisaoCadastralController.saveFiles(acc.Id, 'arquivo2.txt', base64Data);
        Test.stopTest();

        List<ContentVersion> versions = [
            SELECT Id
            FROM   ContentVersion
            WHERE  Title IN ('arquivo1.txt', 'arquivo2.txt')
        ];
        System.assertEquals(2, versions.size(), 'Devem existir dois ContentVersions.');
    }
}