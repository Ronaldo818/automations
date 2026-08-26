public with sharing class B2BAgrosysClienteBatch implements Database.Batchable<SObject>, Database.Stateful, Database.AllowsCallouts {
    
    public integer clienteUnidade;
    Set<String> setCustomerCode = new Set<String>();
    String defaultQuery = 'SELECT Id FROM Account LIMIT 1';
    Map<String, String> mapUnidade = new Map<String, String>();
    Integer maximuAttempts;
    public Integer currentAttempt = 1;
    Boolean success = true;

    public B2BAgrosysClienteBatch(){
        clienteUnidade = 0;
    }
    
    public B2BAgrosysClienteBatch(Integer unidade){
        clienteUnidade = unidade;
    }

    public Database.QueryLocator start(Database.BatchableContext BC) {
        maximuAttempts = Integer.valueOf([SELECT Value__c FROM StandardValue__mdt WHERE DeveloperName = 'getClientMaximumAttempts'].Value__c);
       	return Database.getQueryLocator(defaultQuery);
    }

    public void execute(Database.BatchableContext BC, List<Account> scope) {
        B2BAgrosysCliente integracao = new B2BAgrosysCliente();
        List<B2BAgrosysRetorno.TpCliente> listReturnClientes = new List<B2BAgrosysRetorno.TpCliente>();
        
        B2BAgrosysCliente.FiltroRequisicao filtro = new B2BAgrosysCliente.FiltroRequisicao();
        filtro.unidade = clienteUnidade;

        List<B2BAgrosysRetorno.TpCliente> response = integracao.getCliente(filtro);
        
        if(response == null && currentAttempt < maximuAttempts){
            success = false;
        }
        if(response != null && !response.isEmpty()) {
            listReturnClientes.addAll(response);
        }

        for(Unidade__c unidade : [
            SELECT Id, CodigoUnidade__c
            FROM Unidade__c
        ]) {
            mapUnidade.put(unidade.CodigoUnidade__c, unidade.Id);
        }

        Map<String, List<ClienteUnidade__c>> mapClientexUnidade = new Map<String, List<ClienteUnidade__c>>();
        List<Account> lstClienteUpsert = new  List<Account>();
        List<String> lstClienteKey = new List<String>();
        if(!listReturnClientes.isEmpty()) {
            for(B2BAgrosysRetorno.TpCliente cliente : listReturnClientes) {
                String key = String.valueOf(cliente.cli_codigo);
                
                if(cliente.uni_unidade != null) {
                    String unidadeCliente = String.valueOf(cliente.uni_unidade);

                    ClienteUnidade__c newClienteUnidade = new ClienteUnidade__c();
                    newClienteUnidade.AccountId__r = new Account(ExternalId__c = key);
                    newClienteUnidade.UnidadeId__c = mapUnidade.get(unidadeCliente);
                    newClienteUnidade.ExternalId__c = key + '-' + unidadeCliente;
                    newClienteUnidade.IsActive__c = true;

                    if (!mapClientexUnidade.containsKey(key)) {
                        mapClientexUnidade.put(key, new List<ClienteUnidade__c> {newClienteUnidade});
                    }
                    else {
                        mapClientexUnidade.get(key).add(newClienteUnidade);
                    }
                }

                if(!lstClienteKey.contains(key)) {
                    List<String> emails = new List<String>();
                    if(String.isNotBlank(cliente.cli_email)) {
                        emails = cliente.cli_email.split(';');
                    }

                    Account conta = new Account();
                    if(!emails.isEmpty()) {
                        for(String email : emails) {
                            if(String.isBlank(conta.Emails__c)) {
                                conta.Emails__c = email;
                            }
                            else {
                                conta.Emails__c += '\n' + email;   
                            }
                        }
                    }
                    
                    if(cliente.cli_dtativ != null) {
                        conta.DataAtivacao__c = Date.valueOf(cliente.cli_dtativ);
                    }
                    if(cliente.cli_dtinativ != null) {
                        conta.DataInativacao__c = Date.valueOf(cliente.cli_dtinativ);
                    }
                    if(cliente.cpg_condpag != null) {
                        CondicaoPagamento__c condPagto = new CondicaoPagamento__c();
                        condPagto.CodigoCondicaoPagamento__c = cliente.cpg_condpag;
                        conta.CondicaoPagamento__r = condPagto;
                    }
                    
                    conta.ExternalId__c = key;
                    conta.CodigoClienteERP__c = String.valueOf(cliente.cli_codigo);
                    conta.Empresa__c = cliente.emp_empresa;
                    conta.Unidade__c = cliente.uni_unidade;
                    conta.NomeFantasia__c = cliente.cli_fantasi;
                    conta.RazaoSocial__c = cliente.cli_razsoc;
                    conta.PessoaFisica__c = cliente.cli_fisjur != null ? cliente.cli_fisjur : false;
                    conta.Name = cliente.cli_razsoc;
                    conta.CNPJInput__c = cliente.cli_cnpj;
                    conta.CPFInput__c = cliente.cli_cpf;
                    conta.AplicaST__c = cliente.cli_aplicast != null ? cliente.cli_aplicast : false;
                    conta.TelefoneComercial__c = cliente.cli_foncom;
                    conta.TelefoneFinanceiro__c = cliente.cli_fonfin;
                    conta.TipoCliente__c = String.valueOf(cliente.fxp_tpcli);
                    conta.OverPrice__c = cliente.cli_overprice;
                    conta.VerbaCliente__c = cliente.cli_vlcliverba;
                    conta.LimiteCredito__c = cliente.cli_limcre;
                    conta.Lista__c = cliente.pre_tiporeg;
                    conta.Ocorrencia__c = cliente.pit_tabpre;
                    conta.ClienteEcommerce__c = cliente.cli_ecommerce != null ? cliente.cli_ecommerce : false;
                    conta.DescontoComercial__c = cliente.ped_descon;
                    conta.DescontoFinanceiro__c = cliente.cli_descfin;

                    lstClienteUpsert.add(conta);
                    lstClienteKey.add(key);
                }
            }
        }

        if(!lstClienteUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] results = Database.upsert(lstClienteUpsert, Account.ExternalId__c, false);
            for(Database.UpsertResult sr : results){
                if(!sr.success) {
                    mapClientexUnidade.remove(lstClienteKey[i]);
                    lstClienteKey.remove(i);
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                            NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro do Cliente: ' + lstClienteUpsert[i].CodigoClienteERP__c + '.',
                            ApexClass__c = 'B2BAgrosysClienteBatch',
                            ApexMethod__c = 'execute',
                            AccountId__c = String.isNotBlank(lstClienteUpsert[i].Id) ? lstClienteUpsert[i].Id : null,
                            MessageError__c = 'Status: ' + err.getStatusCode() + '\n\nCampos: ' + err.getFields() + '\n\nMensagem' + err.getMessage(),
                            OperationStatus__c = 'Error');
                            lstLogInsert.add(log);
                    }
                }
                i++;   
            }
            if(!lstLogInsert.isEmpty()) {
                insert lstLogInsert;
            }
        }

        List<ClienteUnidade__c> lstClienteUnidadeUpsert = new List<ClienteUnidade__c>();
        for(List<ClienteUnidade__c> lstAccXUni : mapClientexUnidade.values()) {
            for(ClienteUnidade__c accXUni : lstAccXUni) {
                lstClienteUnidadeUpsert.add(accXUni);
            }
        }
        Database.upsert(lstClienteUnidadeUpsert, ClienteUnidade__c.ExternalId__c,false);
        setCustomerCode.addAll(lstClienteKey);
    }
    
    public void finish(Database.BatchableContext BC) {
        if(!success && currentAttempt < maximuAttempts){           
            B2BAgrosysClienteBatch batch = new B2BAgrosysClienteBatch(clienteUnidade);
            batch.currentAttempt = this.currentAttempt + 1;
            
            LogErro__c log = new LogErro__c();
            //log.Nome__c = 'Get Clients - Extra attempts';
            log.NomeAutomacao__c = 'B2BAgrosysBatch';
            log.TextoErro__c = 'Número de tentativas: ' + currentAttempt;
            insert log;

            Database.executeBatch(batch, 1);
        }
        

        if (!Test.isRunningTest()) {
            B2BAgrosysEnderecoBatch endereco = new B2BAgrosysEnderecoBatch(setCustomerCode);
			Database.executeBatch(endereco, 1); 
        }
    }
}