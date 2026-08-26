public with sharing class RevisaoCadastralController {

    @AuraEnabled
    public static Id saveRevisaoCadastral(String fieldsJson) {
        Map<String, Object> fields = (Map<String, Object>) JSON.deserializeUntyped(fieldsJson);

        RevisaoCadastral__c revisao = new RevisaoCadastral__c();

        if (fields.containsKey('Conta__c') && fields.get('Conta__c') != null) {
            revisao.Conta__c = (Id) fields.get('Conta__c');
        }
        if (fields.containsKey('TipoRevisao__c')) {
            revisao.TipoRevisao__c = (String) fields.get('TipoRevisao__c');
        }
        if (fields.containsKey('Cnpj__c')) {
            String cnpj = (String) fields.get('Cnpj__c');
            revisao.Cnpj__c = cnpj != null ? cnpj.replaceAll('[^A-Za-z0-9]', '').toUpperCase() : null;
        }
        if (fields.containsKey('CPF__c')) {
            String cpf = (String) fields.get('CPF__c');
            revisao.CPF__c = cpf != null ? cpf.replaceAll('[^0-9]', '') : null;
        }
        if (fields.containsKey('Telefone__c')) {
            revisao.Telefone__c = (String) fields.get('Telefone__c');
        }
        if (fields.containsKey('Email__c')) {
            revisao.Email__c = (String) fields.get('Email__c');
        }
        if (fields.containsKey('Comentarios__c')) {
            revisao.Comentarios__c = (String) fields.get('Comentarios__c');
        }

        revisao.Solicitante__c = UserInfo.getUserId();

        insert revisao;
        return revisao.Id;
    }

    // @AuraEnabled(cacheable=true)
    @AuraEnabled
    public static Account findAccountByCnpj(String cnpj) {
        String cleanCnpj = cnpj != null ? cnpj.replaceAll('[^A-Za-z0-9]', '').toUpperCase() : null;
        List<Account> accounts = [
            SELECT Id, Name
            FROM Account
            WHERE CNPJ__c = :cleanCnpj
            LIMIT 1
        ];
        return accounts.isEmpty() ? null : accounts[0];
    }

    // @AuraEnabled(cacheable=true)
    @AuraEnabled
    public static Account findAccountByCpf(String cpf) {
        String cleanCpf = cpf != null ? cpf.replaceAll('[^0-9]', '') : null;
        List<Account> accounts = [
            SELECT Id, Name
            FROM Account
            WHERE CPF__c = :cleanCpf
            LIMIT 1
        ];
        return accounts.isEmpty() ? null : accounts[0];
    }

    @AuraEnabled(cacheable=true)
    public static List<Map<String, String>> getTipoRevisaoOptions() {
        List<Map<String, String>> options = new List<Map<String, String>>();
        Schema.DescribeFieldResult fieldResult = RevisaoCadastral__c.TipoRevisao__c.getDescribe();
        for (Schema.PicklistEntry entry : fieldResult.getPicklistValues()) {
            if (entry.isActive()) {
                options.add(new Map<String, String>{
                    'label' => entry.getLabel(),
                    'value' => entry.getValue()
                });
            }
        }
        return options;
    }

    @AuraEnabled
    public static void saveFiles(Id parentId, String fileName, String base64Data) {
        ContentVersion cv = new ContentVersion();
        cv.Title = fileName;
        cv.PathOnClient = fileName;
        cv.VersionData = EncodingUtil.base64Decode(base64Data);
        insert cv;

        Id contentDocumentId = [
            SELECT ContentDocumentId
            FROM ContentVersion
            WHERE Id = :cv.Id
            LIMIT 1
        ].ContentDocumentId;

        ContentDocumentLink cdl = new ContentDocumentLink();
        cdl.ContentDocumentId = contentDocumentId;
        cdl.LinkedEntityId = parentId;
        cdl.ShareType = 'V';
        insert cdl;
    }
}