public without sharing class B2bGerenciadorPricebook {

    class CustomException extends Exception {}

    public class FlowOutputs {
		@InvocableVariable
		public boolean hasError = false;
		@InvocableVariable
		public String errorMessage = '';
	}

    @InvocableMethod(label='Update Customer Pricebook2')
    public static List<FlowOutputs> sincronizar(List<Account> lstAccount) {
        List<FlowOutputs> result = new List<FlowOutputs>();
        sincronizarAccount(lstAccount[0]);

        for(Integer i = result.size(); i < lstAccount.size(); i++) {
            FlowOutputs newOutput = new FlowOutputs();
            result.add(newOutput);
        }
        return result;
    }

    public static void sincronizarAccount(Account acc) {
        try {
            B2BAgrosysTabelaPreco integracao = new B2BAgrosysTabelaPreco();
            B2BAgrosysRetorno retorno = integracao.getCustomerPricebook2(acc);
    
            if(retorno.hasError) {
                throw new CustomException(retorno.errorMessage);
            }
            else {
                if(retorno.tp_retorno != null && !retorno.tp_retorno.isEmpty() && retorno.tp_retorno[0]?.msg_codigo == 1) {
                    validarPricebook2(retorno.tp_retorno[0]?.msg_retorno[0]?.tp_tabpreco, acc);
                }
                else if(retorno.tp_retorno != null && !retorno.tp_retorno.isEmpty() && retorno.tp_retorno[0]?.msg_codigo != 1) {
                    // 102 = Cliente nÃ£o cadastrado.
                    // 103 = Cliente nÃ£o habilitado no e_commece.
                    if(retorno.tp_retorno[0]?.msg_codigo == 102 || retorno.tp_retorno[0]?.msg_codigo == 103) {
                        delete [
                            SELECT Id
                            FROM BuyerAccount
                            WHERE BuyerId = :acc.Id
                        ];

                        acc.ClienteEcommerce__c = false;
                        update acc;
                    }
                    else {
                        throw new CustomException('Ocorreu um erro, por favor tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema junto com o ID do erro: ' + retorno.logId + '\n\nCausa do Erro: ' + retorno.tp_retorno[0]?.msg_descricao);
                    }
                }
            }
        } catch (Exception ex) {
            User us = [
                SELECT Id, Name
                FROM User
                WHERE Id = :System.UserInfo.getUserId()
            ];

            ErrorLog__e errorLog = new ErrorLog__e();
            errorLog.AutomationType__c = 'Apex';
            errorLog.AutomationName__c = 'B2bGerenciadorPricebook';
            errorLog.Status__c = 'Erro';
            errorLog.Name__c = 'ERROR B2bGerenciadorPricebook';
            errorLog.ErrorMessage__c = '';
            
            if(us != null) {
                errorLog.ErrorMessage__c += 'User: ' + us.Name + ' - ' + us.Id + '\n\n';
            }

            if(acc != null) {
                errorLog.Name__c += ' - ' + acc.CodigoClienteERP__c;
                errorLog.ErrorMessage__c += 'Account: ' + acc.Name + ' - ' + acc.Id + '\n\n';
            }

            errorLog.ErrorMessage__c += 'Método: sincronizar' + '\n\nType: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();

            String eventUUID;
            Database.SaveResult sr = EventBus.publish(errorLog);
            if (sr.isSuccess()) {
                eventUUID = EventBus.getOperationId(sr);
            } else {
                for(Database.Error err : sr.getErrors()) {
                    System.debug('Error returned: ' + err.getStatusCode() + ' - ' + err.getMessage());
                }
            }
            throw new CustomException('Ocorreu um erro, por favor tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema junto com o UUID do erro: ' + eventUUID);
        }
    }

    private static void validarPricebook2(List<B2BAgrosysRetorno.TabPreco> lstTabPreco, Account account) {
        // Criar Pricebook2, PricebookEntry, BuyerGroup e CommerceEntitlementPolicy
        if(!lstTabPreco.isEmpty()) {
            //Lista todos os SKU's e chaves de pricebook
            Id standardPBId = Test.isRunningTest() ? Test.getStandardPricebookId() : [SELECT Id FROM Pricebook2 WHERE IsStandard = true LIMIT 1].Id;
            Set<String> setBGKey = new Set<String>();
            Set<String> setPricebookKey = new Set<String>();
            Set<String> setProductSKU = new Set<String>();
            Set<String> setPBEKey = new Set<String>();
            Set<String> setUnidadeProdutoKey = new Set<String>();
            Set<String> setPolicyKey = new Set<String>();
            
            for(B2BAgrosysRetorno.TabPreco item : lstTabPreco) {
                String key = String.valueOf(item.cli_codigo);
                // String key = item.emp_empresa + '-' + item.uni_unidade + '-' + item.pit_tabpre + '-' + item.cli_codigo;
                setBGKey.add(key);
                setPricebookKey.add(key);
                setPBEKey.add(key + '-' + item.pro_codigo);
                setProductSKU.add(String.valueof(item.pro_codigo));
                setUnidadeProdutoKey.add(item.uni_unidade + '-' + item.pro_codigo);
                setPolicyKey.add(key + '-' + item.pro_codigo);
            }

            Set<String> setPBEStandard = new Set<String>() ;
            for(PriceBookEntry priceStandard : [
                SELECT Id, Product2.CodigoProdutoERP__c
                FROM PricebookEntry
                WHERE Product2.CodigoProdutoERP__c IN :setProductSKU
                AND  PriceBook2Id = :standardPBId
            ]){
                setPBEStandard.add(priceStandard.Product2.CodigoProdutoERP__c);
            }
                
            //Mapear produtos de acordo com a chave
            Map<String, String> mapSKUxProductId = new Map<String, String>();
            Map<String, Decimal> mapSKUxPesoUnitario = new Map<String, Decimal>();

            for(Product2 product : [
                SELECT Id, CodigoProdutoERP__c, PesoUnitario__c
                FROM Product2
                WHERE CodigoProdutoERP__c IN :setProductSKU
            ]) {
                mapSKUxProductId.put(product.CodigoProdutoERP__c, product.Id);
                mapSKUxPesoUnitario.put(product.CodigoProdutoERP__c, product.PesoUnitario__c);
            }

            //Mapear pricebook de acordo com a chave
            Map<String, String> mapKeyxPriceBookId = new Map<String, String>();
            for(Pricebook2 pricebook : [
                SELECT Id, KeyBuyerGroup__c
                FROM Pricebook2
                WHERE KeyBuyerGroup__c IN :setPricebookKey
                AND Id != :standardPBId
            ]) {
                mapKeyxPriceBookId.put(pricebook.KeyBuyerGroup__c, pricebook.Id);
                setPricebookKey.remove(pricebook.KeyBuyerGroup__c);
            }

            //Mapear PricebookEntry existentes
            Map<String, PricebookEntry> mapKeyPBExPBE = new Map<String, PricebookEntry>();

            for(PricebookEntry pbe : [
                SELECT Id, KeyBuyerGroup__c
                FROM PricebookEntry
                WHERE KeyBuyerGroup__c IN :setPBEKey
            ]) {
                mapKeyPBExPBE.put(pbe.KeyBuyerGroup__c, pbe);
            }

            //Mapear unidades do produto existente
            Map<String, Boolean> mapUnidadeProdutoKeyxTipoMedida = new Map<String, Boolean>();

            for(UnidadeProduto__c unidade : [
                SELECT Key__c, MultiplicarQuantidePorPeso__c
                FROM UnidadeProduto__c
                WHERE Key__c IN :setUnidadeProdutoKey
            ]) {
                mapUnidadeProdutoKeyxTipoMedida.put(unidade.Key__c, unidade.MultiplicarQuantidePorPeso__c);
            }
            
            Map<String, String> mapKeyxNewPriceBookId = new Map<String, String>();
            List<BuyerAccount> lstBuyerAccountUpdate = new List<BuyerAccount>();
            //Verifica se existe pricebooks pendentes para criar
            if (!setPricebookKey.isEmpty()) {
                //Listar e Criar os Pricebook2 pendentes
                List<Pricebook2> listNewPriceBook = new List<Pricebook2>();
                for(String key : setPricebookKey) {
                    Pricebook2 newPricebook = new Pricebook2();
                    newPricebook.CodigoClienteERP__c    = Integer.valueOf(key);
                    newPricebook.AccountId__c           = account.Id;
                    newPricebook.Name                   = account.Name + ' - ' + key;
                    newPricebook.IsActive               = true;
                    listNewPriceBook.add(newPricebook);
                    setPricebookKey.remove(key);
                }
                insert listNewPriceBook;

                //Criar os BuyerGroup que devem ser associados a esses Pricebook2 e mapear os novos Pricebook2
                List<BuyerGroup> listNewBuyerGroup = new List<BuyerGroup>();
                List<CommerceEntitlementPolicy> listNewCommerceEntitlementPolicy = new List<CommerceEntitlementPolicy>();

                for (Pricebook2 pricebook : listNewPriceBook) {
                    String key = String.valueOf(pricebook.CodigoClienteERP__c);
                    mapKeyxNewPriceBookId.put(key, pricebook.Id);

                    BuyerGroup newBuyerGroup = new BuyerGroup();
                    newBuyerGroup.Name = pricebook.Name;
                    newBuyerGroup.Key__c = key;
                    listNewBuyerGroup.add(newBuyerGroup);

                    CommerceEntitlementPolicy newCommerceEntitlementPolicy = new CommerceEntitlementPolicy();
                    newCommerceEntitlementPolicy.Name = pricebook.Name;
                    newCommerceEntitlementPolicy.IsActive = true;
                    newCommerceEntitlementPolicy.CanViewPrice = true;
                    newCommerceEntitlementPolicy.CanViewProduct = true;
                    newCommerceEntitlementPolicy.Key__c = key;
                    listNewCommerceEntitlementPolicy.add(newCommerceEntitlementPolicy);

                }
                insert listNewBuyerGroup;
                upsert listNewCommerceEntitlementPolicy Key__c;
                
                //Cria o link entre o BuyerGroup, Pricebook2 e WebStore (loja). Logo após mapear os BuyerGroup ID pelo nome
                WebStore store = [SELECT Id FROM WebStore WHERE Name = 'Avivar LWR'];
                List<BuyerGroupPricebook> listNewBuyerGroupPricebook = new List<BuyerGroupPricebook>();
                List<WebStoreBuyerGroup> listNewWebStoreBuyerGroup = new List<WebStoreBuyerGroup>();
                Map<String, String> mapBuyerGroupNamexId = new Map<String, String>();
                List<BuyerGroupMember> listBuyerGroupMember = new List<BuyerGroupMember>();
                
                for(BuyerGroup buyerGroup : listNewBuyerGroup) {
                    BuyerGroupPricebook newBuyerGroupPricebook = new BuyerGroupPricebook();
                    newBuyerGroupPricebook.BuyerGroupId = buyerGroup.Id;
                    newBuyerGroupPricebook.Pricebook2Id = mapKeyxNewPriceBookId.get(buyerGroup.Key__c);
                    newBuyerGroupPricebook.IsActive = true;
                    listNewBuyerGroupPricebook.add(newBuyerGroupPricebook);
                    
                    WebStoreBuyerGroup storeLink = new WebStoreBuyerGroup();
                    storeLink.WebStoreId = store.Id;
                    storeLink.BuyergroupId = buyerGroup.Id;
                    listNewWebStoreBuyerGroup.add(storeLink);
                    
                    mapBuyerGroupNamexId.put(buyerGroup.Name, buyerGroup.Id);

                    BuyerGroupMember newBuyerGroupMember = new BuyerGroupMember();
                    newBuyerGroupMember.BuyerGroupId = buyerGroup.Id;
                    newBuyerGroupMember.BuyerId = account.Id;
                    listBuyerGroupMember.add(newBuyerGroupMember);
                }
                insert listNewBuyerGroupPricebook;
                insert listNewWebStoreBuyerGroup;
				insert listBuyerGroupMember;

                //Cria o link entre o buyer group e policy
                List<CommerceEntitlementBuyerGroup> listNewPolicyBuyerGroup = new List<CommerceEntitlementBuyerGroup>();

                for(CommerceEntitlementPolicy policy : listNewCommerceEntitlementPolicy) {
                    CommerceEntitlementBuyerGroup newCommerceEntitlementBuyerGroup = new CommerceEntitlementBuyerGroup();
                    newCommerceEntitlementBuyerGroup.BuyerGroupId = mapBuyerGroupNamexId.get(policy.Name);
                    newCommerceEntitlementBuyerGroup.PolicyId = policy.Id;
                    listNewPolicyBuyerGroup.add(newCommerceEntitlementBuyerGroup);
                }
                insert listNewPolicyBuyerGroup;

                BuyerAccount buyerAcc = [
                    SELECT Id, PendingEcommerceRecords__c 
                    FROM BuyerAccount
                    WHERE BuyerId = :account.Id
                ];

                if(buyerAcc.PendingEcommerceRecords__c) {
                    buyerAcc.PendingEcommerceRecords__c = false;
                    lstBuyerAccountUpdate.add(buyerAcc);
                } 
            }

            //Mapear CommerceEntitlementProduct existentes
            Map<String, CommerceEntitlementProduct> mapKeyxCEP = new Map<String, CommerceEntitlementProduct>();
            for(CommerceEntitlementProduct entitlementProduct : [
                SELECT Id, PolicyId, ProductId
                FROM CommerceEntitlementProduct
                WHERE Key__c IN :setPolicyKey
            ]) {
                String key = String.valueOf(entitlementProduct.PolicyId) + String.valueOf(entitlementProduct.ProductId);
                mapKeyxCEP.put(key, entitlementProduct);
            }

            //Mapear pbes existentes
            Map<String, String> mapKeyPBExPolicyId = new Map<String, String>();
            for(CommerceEntitlementBuyerGroup EBG : [
                SELECT BuyerGroup.Key__c, PolicyId
                FROM CommerceEntitlementBuyerGroup
                WHERE BuyerGroup.Key__c IN :setBGKey
            ]) {
                mapKeyPBExPolicyId.put(EBG.BuyerGroup.Key__c, EBG.PolicyId);
            }

            List<PricebookEntry> listNewPBEStandard = new List<PricebookEntry>();
            List<PricebookEntry> listNewPBE = new List<PricebookEntry>();
            List<PricebookEntry> listUpdatePBE = new List<PricebookEntry>();
            List<CommerceEntitlementProduct> listNewCEP = new List<CommerceEntitlementProduct>();
            List<CommerceEntitlementProduct> listDeleteCEP = new List<CommerceEntitlementProduct>();
            

            for(B2BAgrosysRetorno.TabPreco item : lstTabPreco) {
                String productCode = String.valueOf(item.pro_codigo);
                if (mapSKUxProductId.containsKey(productCode)) {
                    String keyPricebook = String.valueOf(item.cli_codigo);
                    // String keyPricebook = item.emp_empresa + '-' + item.uni_unidade + '-' + item.pit_tabpre + '-' + item.cli_codigo;
                    String productId = mapSKUxProductId.get(productCode);
                    String policyId = mapKeyPBExPolicyId.get(keyPricebook);
                    String keyPBE = keyPricebook + '-' + productCode;
                    String keyCEP = policyId + productId;
                    String unidadeProdutoKey = item.uni_unidade + '-' + productCode;

                    PricebookEntry newPBE = mapKeyPBExPBE.containsKey(keyPBE) ? mapKeyPBExPBE.get(keyPBE) : new PricebookEntry();
                    newPBE.KeyBuyerGroup__c = keyPBE;
                    newPBE.Unidade__c = String.valueOf(item.uni_unidade);

                    if(mapUnidadeProdutoKeyxTipoMedida.containsKey(unidadeProdutoKey)){
                        newPBE.UnitPrice = (item.pre_tabela * mapSKUxPesoUnitario.get(productCode)).setScale(2);
                    }
                    else {
                        newPBE.UnitPrice = item.pre_tabela;
                    }

                    if (newPBE.UnitPrice > 0) {
                        newPBE.IsActive = true;
                        if(!mapKeyxCEP.containsKey(keyCEP)) {
                            CommerceEntitlementProduct link = new CommerceEntitlementProduct();
                            link.PolicyId = policyId;
                            link.ProductId = productId;
                            link.key__c = keyPricebook + '-' + productCode;
                            listNewCEP.add(link);
                        }
                    }
                    else {
                        newPBE.IsActive = false;
                        if(mapKeyxCEP.containsKey(keyCEP)) {
                            listDeleteCEP.add(mapKeyxCEP.get(keyCEP));
                        }
                    }

                    if(mapKeyxNewPriceBookId.containsKey(keyPricebook) && String.isBlank(newPBE.Id)){
                        newPBE.Product2Id = productId;
                        newPBE.Pricebook2Id = mapKeyxNewPriceBookId.get(keyPricebook);
                        listNewPBE.add(newPBE);
                    }
                    else if(mapKeyxPriceBookId.containsKey(keyPricebook) && String.isBlank(newPBE.Id)){
                        newPBE.Product2Id = productId;
                        newPBE.Pricebook2Id = mapKeyxPriceBookId.get(keyPricebook);
                        listNewPBE.add(newPBE);
                    }
                    else {
                        listUpdatePBE.add(newPBE);
                    }

                    if(!setPBEStandard.contains(productCode)) {
						PricebookEntry newStandardPrice = new PricebookEntry(
							Pricebook2Id = standardPBId,
							Product2Id = productId,
							UnitPrice = 0,
							IsActive = true
                        ); 
                        listNewPBEStandard.add(newStandardPrice);   
                    }
                }
            }

            if (!listNewPBEStandard.isEmpty()) {
                List<Database.SaveResult> rs = Database.insert(listNewPBEStandard, false);
                createLogOnFailed('Standard Price Insert - ' + account.CodigoClienteERP__c, rs);
            }

            if (!listNewPBE.isEmpty()) {
                List<Database.SaveResult> rs = Database.insert(listNewPBE, false);
                createLogOnFailed('Price Insert - ' + account.CodigoClienteERP__c, rs);
            }

            if (!listUpdatePBE.isEmpty()) {
                List<Database.SaveResult> rs = Database.update(listUpdatePBE, false);
                createLogOnFailed('Price update - ' + account.CodigoClienteERP__c, rs);
            }

            if (!listNewCEP.isEmpty()) {
                List<Database.SaveResult> rs = Database.insert(listNewCEP, false);
                createLogOnFailed('Entitlement insert - ' + account.CodigoClienteERP__c, rs);
            }

            if (!listDeleteCEP.isEmpty()) {
                delete listDeleteCEP;
            }
            
            if (!lstBuyerAccountUpdate.isEmpty()) {
                update lstBuyerAccountUpdate;
            }
        }
    }

    public static void createLogOnFailed(String operationName, List<Database.SaveResult> results){
        LogErro__c l = new LogErro__c();
        Integer errorCount = 0;
        l.Nome__c = operationName + Datetime.now().toString();
        l.TextoErro__c = '';
        l.Status__c = 'Erro';
        l.TipoAutomacao__c = 'Apex';
        l.NomeAutomacao__c = 'B2bGerenciadorPricebook';
        for(Database.SaveResult r : results){
            if(!r.success){
                l.TextoErro__c += r.toString() + '\n';
                errorCount++;
            }
        }
        if(!String.isBlank(l.TextoErro__c)){
            l.TextoErro__c = 'Sucesso: ' + String.valueOf(results.size() - errorCount) + '/' + results.size() + '\n' + l.TextoErro__c.left(19900);
            insert l;
        }
    }
}