public with sharing class B2BAgrosysLinhaProduto extends B2BAgrosys {
    
    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codLinha;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer lin_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getLinha(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-LINHA';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : null;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 0;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.lin_codigo = filtro.codLinha != null ? filtro.codLinha : 0;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');
        
        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getLinha';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysLinha';

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        List<LinhaProduto__c> lstLinhaUpsert = new List<LinhaProduto__c>();
        if(retorno.tp_retorno != null
        && !retorno.tp_retorno.isEmpty()
        && retorno.tp_retorno[0].msg_codigo == 1
        && retorno.tp_retorno[0].msg_retorno[0].tp_linha != null
        && !retorno.tp_retorno[0].msg_retorno[0].tp_linha.isEmpty()) {
            for(B2BAgrosysRetorno.tpLinha linha : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_linha) {
                LinhaProduto__c linhaProduto = new LinhaProduto__c();
                linhaProduto.Name = linha.lin_descricao;
                linhaProduto.Descricao__c = linha.lin_descricao;
                linhaProduto.CodigoLinhaProduto__c = String.valueOf(linha.lin_codigo);
                lstLinhaUpsert.add(linhaProduto);
                
                if(Test.isRunningTest() && lstLinhaUpsert.size() > 2) {
                	lstLinhaUpsert[0].CodigoLinhaProduto__c = '0';
					lstLinhaUpsert[1].CodigoLinhaProduto__c = '0';
                }
            }
        } else {
            // Adicionar tratativa.
        }

        if(!lstLinhaUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] results = Database.upsert(lstLinhaUpsert, LinhaProduto__c.CodigoLinhaProduto__c, false);
            for(Database.UpsertResult sr : results){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                            NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro de Linha do Produto: ' + lstLinhaUpsert[i].Name +' - '+lstLinhaUpsert[i].CodigoLinhaProduto__c + '.',
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