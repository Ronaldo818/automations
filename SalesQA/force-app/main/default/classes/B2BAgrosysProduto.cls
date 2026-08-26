public with sharing class B2BAgrosysProduto extends B2BAgrosys {

    public class FiltroRequisicao {
        public Integer empresa;
        public Integer unidade;
        public Boolean indUnderline;
        public Integer codigoProduto;
        public String nomeProduto;
        public Boolean situacaoProduto;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer pro_codigo;
        public String pro_descricao;
        public Boolean pro_situa;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getProduto(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-PRODUTO';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new List<TpEnvio>();

        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.pro_codigo = filtro.codigoProduto != null ? filtro.codigoProduto : 0;
        tpEnvio.pro_descricao = filtro.nomeProduto != null ? filtro.nomeProduto : '';
        tpEnvio.pro_situa = filtro.situacaoProduto != null ? filtro.situacaoProduto : true;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getProduto';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysProduto';

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        List<Product2> lstProdutoUpsert = new List<Product2>();
        List<UnidadeProduto__c> lstUnidadeProduto = new List<UnidadeProduto__c>();
        List<String> listProductsSKU = new List<String>();
        if(retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0]?.msg_codigo == 1
            && retorno.tp_retorno[0]?.msg_retorno[0]?.tp_produto != null
            && !retorno.tp_retorno[0]?.msg_retorno[0]?.tp_produto.isEmpty()) {
                for (B2BAgrosysRetorno.TpProduto produto : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_produto) {
                    String sku = String.valueOf(produto.pro_codigo);
                    if (!listProductsSKU.contains(sku)) {
                        listProductsSKU.add(sku);
                        Product2 produ = new Product2();
                        produ.CodigoProdutoERP__c = sku;
                        produ.Name = String.isBlank(produ.Name) ? produto.pro_descricao : produ.Name;
                        produ.ProductCode = sku;
                        produ.StockKeepingUnit = sku;
                        produ.Empresa__c = produto.emp_empresa;
                        //produ.Unidade__c = produto.uni_unidade;
                        produ.UnidadeMedida__c = produto.med_codigo;
                        produ.PesoUnitario__c = produto.pro_psunitario;
                        produ.IsActive = produto.pro_situa;
                        produ.PesoEmbalagem__c = produto.emb_peso;
                        produ.AlturaEmbalagem__c = produto.emb_altura;
                        produ.LarguraEmbalagem__c = produto.emb_largura;
                        produ.ComprimentoEmbalagem__c = produto.emb_comprimento;
                        if(produto.fam_codigo != null && produto.fam_codigo > 0) {
                            FamiliaProduto__c familia = new FamiliaProduto__c();
                            familia.CodigoFamilia__c = String.valueOf(produto.fam_codigo);
                            produ.FamiliaId__r = familia;
                        }
                        if(produto.gru_codigo != null && produto.gru_codigo > 0) {
                            GrupoProduto__c grupo = new GrupoProduto__c();
                            grupo.CodigoGrupo__c = String.valueOf(produto.gru_codigo);
                            produ.GrupoProdutoId__r = grupo;
                        }
                        if(produto.lin_codigo != null && produto.lin_codigo > 0) {
                            LinhaProduto__c linha = new LinhaProduto__c();
                            linha.CodigoLinhaProduto__c = String.valueOf(produto.lin_codigo);
                            produ.LinhaProduto__r = linha;
                        }
                        if(produto.ori_codigo != null && produto.ori_codigo > 0) {
                            OrigemProduto__c origem = new OrigemProduto__c();
                            origem.CodigoOrigemProduto__c = String.valueOf(produto.ori_codigo);
                            produ.OrigemProduto__r = origem;
                        }
                        lstProdutoUpsert.add(produ);
                    }
                    Product2 product = new Product2();
                    product.CodigoProdutoERP__c                 = sku;
                    UnidadeMedida__c unidadePedido              = new UnidadeMedida__c();
                    unidadePedido.CodigoUnidadeMedida__c        = produto.pro_uniped;
                    UnidadeMedida__c unidadeFaturamento         = new UnidadeMedida__c();
                    unidadeFaturamento.CodigoUnidadeMedida__c   = produto.pro_unifat;
                    UnidadeProduto__c unidade                   = new UnidadeProduto__c();
                    unidade.Key__c                              = String.valueOf(produto.uni_unidade)+'-'+sku;
                    unidade.Unidade__c                          = produto.uni_unidade;
                    unidade.Product__r                          = product;
                    unidade.UnidadePedido__r                    = unidadePedido;
                    unidade.UnidadeFaturamento__r               = unidadeFaturamento;
                    lstUnidadeProduto.add(unidade);
                }

            if(!lstProdutoUpsert.isEmpty()) {
                List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
                Integer i = 0;
                Database.UpsertResult[] resultsProducts = Database.upsert(lstProdutoUpsert, Product2.CodigoProdutoERP__c, false);
                for(Database.UpsertResult sr : resultsProducts){
                    if(!sr.success) {
                        for(Database.Error err : sr.getErrors()) {
                            IntegrationLog__c log = new IntegrationLog__c(
                                NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro do Produto: '+lstProdutoUpsert[i].CodigoProdutoERP__c+'.',
                                MessageError__c = 'Status: ' + err.getStatusCode() + '\n\nCampos: ' + err.getFields() + '\n\nMensagem' + err.getMessage(),
                                OperationStatus__c = 'Error');
                                lstLogInsert.add(log);
                        }
                    }
                    i++;   
                }

                i = 0;
                Database.UpsertResult[] resultsUnidades = Database.upsert(lstUnidadeProduto, UnidadeProduto__c.Key__c, false);
                for(Database.UpsertResult sr : resultsUnidades){
                    if(!sr.success) {
                        for(Database.Error err : sr.getErrors()) {
                            IntegrationLog__c log = new IntegrationLog__c(
                                NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro de Unidade do Produto: ' + lstUnidadeProduto[i].Key__c + '.',
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
        }
        
        //Criar Price Book Entry standard para os novos produtos
        List<PricebookEntry> lstNewPBE = new List<PricebookEntry>();
        String standardPBId = Test.isRunningTest() ? Test.getStandardPricebookId() : [SELECT Id FROM Pricebook2 WHERE IsStandard = true LIMIT 1].Id;
        for (Product2 product : [
            SELECT Id
            FROM Product2
            WHERE Id NOT IN (
                SELECT Product2Id
                FROM PricebookEntry
                WHERE Pricebook2Id = :standardPBId
            )
        ]) {
            PricebookEntry newPBE = new PricebookEntry(
                Pricebook2Id = standardPBId,
				Product2Id = product.Id,
				UnitPrice = 0,
				IsActive = true
            );
            lstNewPBE.add(newPBE);
        }
        if (!lstNewPBE.isEmpty()) insert lstNewPBE;
    }

}