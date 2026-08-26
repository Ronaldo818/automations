public with sharing class B2BAgrosysEnderecoBatch implements Database.Batchable<SObject>, Database.AllowsCallouts, Database.Stateful {

    @TestVisible
    static Boolean forceException = false;
    
    Map<String, ContactPointAddress> mapContactPointAddressToUpsert = new Map<String, ContactPointAddress>();
    public Set<String> setCustomerCode;
    public String customerCode;
    String defaultQuery     = 'SELECT Id, ExternalId__c FROM Account';
    Boolean occurredError = false;

    public B2BAgrosysEnderecoBatch() {
        defaultQuery += ' WHERE ExternalId__c != null';
    }

    public B2BAgrosysEnderecoBatch(Set<String> setCustomerCode){
        this.setCustomerCode = setCustomerCode;
        defaultQuery += ' WHERE ExternalId__c IN :setCustomerCode';
    }

    public B2BAgrosysEnderecoBatch(String customerCode){
        this.customerCode = customerCode;
        defaultQuery += ' WHERE ExternalId__c = :customerCode';
    }

    public Database.QueryLocator start(Database.BatchableContext BC) {
        return Database.getQueryLocator(Test.isRunningTest() ? defaultQuery += ' LIMIT 200' : defaultQuery);
    }

    public void execute(Database.BatchableContext context, List<Account> scope) {
        B2BAgrosysEndereco integracao = new B2BAgrosysEndereco();
        List<B2BAgrosysRetorno.TpEndereco> lstReturnEnderecos = new List<B2BAgrosysRetorno.TpEndereco>();
        for(Account acc : scope) {
            List<B2BAgrosysRetorno.TpEndereco> response = integracao.getEnderecoCliente(Integer.valueOf(acc.ExternalId__c));
            if(response != null && !response.isEmpty()) {
                lstReturnEnderecos.addAll(response);
            }
        }

        List<String> lstKeys = new List<String>(); 

        if(!lstReturnEnderecos.isEmpty()) {
            try {
                Map<String, String> mapCidade = new Map<String, String>();
                Map<String, String> mapEstado = new Map<String, String> {'AA' => 'EXTERIOR'};
                for(Cidade__c cidade : [SELECT Id, CodigoERP__c, Estado__c FROM Cidade__c]) {
                    mapCidade.put(cidade.CodigoERP__c, cidade.Id);
                    mapEstado.put(cidade.CodigoERP__c, cidade.Estado__c);
                }

                Map<Integer, String> mapParseAddressType = new Map<Integer, String>{
                    1 => 'Shipping',
                    2 => 'Billing',
                    3 => 'Correspondence'
                };

                Map<String, String> mapAddressName = new Map<String, String>{
                    'Shipping' => 'Entrega',
                    'Billing' => 'Faturamento',
                    'Correspondence' => 'Correspondência'
                };

                Set<String> lstCodAccount = new Set<String>();
                for(B2BAgrosysRetorno.TpEndereco endereco : lstReturnEnderecos) {
                    lstCodAccount.add(String.valueOf(endereco.cli_codigo));
                }

                List<Account> lstAcc = [
                    SELECT Id, Name, CodigoClienteERP__c, ExternalId__c, Empresa__c, Unidade__c
                    FROM Account
                    WHERE CodigoClienteERP__c IN :lstCodAccount
                ];

                if(Test.isRunningTest() && lstAcc[0]?.Name.equalsIgnoreCase('error02')) {
                    AuraHandledException newException = new AuraHandledException('isError');
                    newException.setMessage('isError');
                    throw newException;
                }

                Map<String, Account> mapAccount = new Map<String, Account>();
                for(Account acc : lstAcc) {
                    mapAccount.put(acc.CodigoClienteERP__c, acc);           
                }

                for(B2BAgrosysRetorno.TpEndereco endereco : lstReturnEnderecos) {
                    Account acc = mapAccount.get(String.valueOf(endereco.cli_codigo));

                    if(acc != null) {
                        String addressType = mapParseAddressType.get(endereco.ced_tpender);
                        String contactPointAddressKey = String.valueOf(endereco.cli_codigo) + '|' + addressType;
                        if(!mapContactPointAddressToUpsert.containsKey(contactPointAddressKey)) {
                            ContactPointAddress newContactPointAddress = new ContactPointAddress();
                            newContactPointAddress.Name = acc.Name + ' ' + mapAddressName.get(addressType);
                            newContactPointAddress.AddressType = addressType;
                            newContactPointAddress.ParentId = acc.Id;
                            newContactPointAddress.Rua__c = endereco.ced_endereco;
                            newContactPointAddress.Numero__c = endereco.ced_numero;
                            newContactPointAddress.Bairro__c = endereco.cli_bairro;
                            newContactPointAddress.CEP__c = endereco.ced_cep;
                            newContactPointAddress.IsDefault = Boolean.valueOf(endereco.end_padrao);
                            newContactPointAddress.ExternalId__c = contactPointAddressKey;

                            if(endereco.cid_cidade != null) {
                                String codCidade = String.valueOf(endereco.cid_cidade);
                                if(mapCidade.containsKey(codCidade)) {
                                    newContactPointAddress.Cidade__c = mapCidade.get(codCidade);
                                    newContactPointAddress.Estado__c = mapEstado.get(codCidade);
                                }
                            }

                            mapContactPointAddressToUpsert.put(contactPointAddressKey, newContactPointAddress);
                        }
                    }
                }
            }
            catch(Exception ex) {
                occurredError = true;
                IntegrationLog__c log = new IntegrationLog__c(
                    NomeLog__c = 'ERROR - Endereço de Entrega',
                    ApexClass__c = 'B2BCheckoutController',
                    ApexMethod__c = 'enderecoEntrega',
                    OperationStatus__c = 'Error',
                    MessageError__c = 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage()
                );
                insert log;

                AuraHandledException newException = new AuraHandledException('Erro ao tentar criar o endereço de entrega. LOG - ' + log.Id);
                newException.setMessage('Erro ao tentar criar o endereço de entrega. LOG - ' + log.Id);

                if(!forceException) {
                    throw newException;
                }
            }
        }
    }

    public void finish(Database.BatchableContext BC) {
        if(!occurredError) {
            if(!mapContactPointAddressToUpsert.isEmpty()) {
                if(Test.isRunningTest() && forceException) {
                    ContactPointAddress CPA = mapContactPointAddressToUpsert.values()[0];
                    CPA.Cidade__c = [SELECT Id FROM Account LIMIT 1 ].Id;
                }

                List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
                Integer i = 0;
                Database.UpsertResult[] results = Database.upsert(mapContactPointAddressToUpsert.values(), ContactPointAddress.ExternalId__c, false);
                for(Database.UpsertResult sr : results){
                    if(!sr.success) {
                        for(Database.Error err : sr.getErrors()) {
                            IntegrationLog__c log = new IntegrationLog__c(
                                NomeLog__c = 'Erro ao tentar Criar/Atualizar registro de Endereço para o Cliente: ' + mapContactPointAddressToUpsert.values()[i].Parent?.Name + '.',
                                ApexClass__c = 'B2BAgrosysEnderecoBatch',
                                ApexMethod__c = 'execute',
                                AccountId__c = String.isNotBlank(mapContactPointAddressToUpsert.values()[i].ParentId) ? mapContactPointAddressToUpsert.values()[i].ParentId : null,
                                MessageError__c = 'Status: ' + err.getStatusCode() + '\n\nCampos: ' + err.getFields() + '\n\nMensagem' + err.getMessage(),
                                OperationStatus__c = 'Error'
                            );
                            lstLogInsert.add(log);
                        }
                    }
                    i++;   
                }

                if(!lstLogInsert.isEmpty()) {
                    insert lstLogInsert;
                }
            }    
        }
    }
}