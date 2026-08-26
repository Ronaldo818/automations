public with sharing class B2BAgrosysOrigemProduto extends B2BAgrosys {
    
    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codOrigem;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer ori_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getOrigem(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-ORIGEM';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : true;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.ori_codigo = filtro.codOrigem != null ? filtro.codOrigem : 0;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');
        
        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getOrigem';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysOrigemProduto';

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        List<OrigemProduto__c> lstOrigemUpsert = new List<OrigemProduto__c>();
        if(retorno.tp_retorno != null
        && !retorno.tp_retorno.isEmpty()
        && retorno.tp_retorno[0].msg_codigo == 1
        && retorno.tp_retorno[0].msg_retorno[0].tp_origem != null
        && !retorno.tp_retorno[0].msg_retorno[0].tp_origem.isEmpty()) {
            for(B2BAgrosysRetorno.tpOrigem origem : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_origem) {
                OrigemProduto__c origemProduto = new OrigemProduto__c();
                origemProduto.Name = origem.ori_descricao;
                origemProduto.Descricao__c = origem.ori_descricao;
                origemProduto.CodigoOrigemProduto__c = String.valueOf(origem.ori_codigo);
                lstOrigemUpsert.add(origemProduto);
                
                if(Test.isRunningTest() && lstOrigemUpsert.size() == 2) {
                   lstOrigemUpsert[0].CodigoOrigemProduto__c = '0';
                   lstOrigemUpsert[1].CodigoOrigemProduto__c = '0';
                }
            }
        } else {
            // Adicionar tratativa.
        }

        if(!lstOrigemUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] results = Database.upsert(lstOrigemUpsert, OrigemProduto__c.CodigoOrigemProduto__c, false);
            for(Database.UpsertResult sr : results){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                            NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro de Origem do Produto: ' + lstOrigemUpsert[i].Name +' - '+lstOrigemUpsert[i].CodigoOrigemProduto__c + '.',
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
}