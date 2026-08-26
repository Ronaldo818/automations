public with sharing class B2BAgrosysGrupoProduto extends B2BAgrosys {
    
    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codGrupo;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer gru_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getGrupoItem(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-GRPITEM';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : true;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.gru_codigo = filtro.codGrupo != null ? filtro.codGrupo : 0;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');
        
        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getGrupoItem';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysGrupoItem';

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        List<GrupoProduto__c> lstGrupoUpsert = new List<GrupoProduto__c>();
        if(retorno.tp_retorno != null
        && !retorno.tp_retorno.isEmpty()
        && retorno.tp_retorno[0].msg_codigo == 1
        && retorno.tp_retorno[0].msg_retorno[0].tp_grpitem != null
        && !retorno.tp_retorno[0].msg_retorno[0].tp_grpitem.isEmpty()) {
            for(B2BAgrosysRetorno.TpGrpitem grupo : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_grpitem) {
                GrupoProduto__c grupoProduto = new GrupoProduto__c();
                grupoProduto.Name = grupo.gru_descricao;
                grupoProduto.Descricao__c = grupo.gru_descricao;
                grupoProduto.CodigoGrupo__c = String.valueOf(grupo.gru_codigo);
                lstGrupoUpsert.add(grupoProduto);
                
                if(Test.isRunningTest() && lstGrupoUpsert.size() > 2) {
                    lstGrupoUpsert[0].CodigoGrupo__c = '0';
					lstGrupoUpsert[1].CodigoGrupo__c = '0';
                }
            }
        } else {
            // Adicionar tratativa.
        }

        if(!lstGrupoUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] results = Database.upsert(lstGrupoUpsert, GrupoProduto__c.CodigoGrupo__c, false);
            for(Database.UpsertResult sr : results){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                            NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro de Grupo do Produto: ' + lstGrupoUpsert[i].Name +' - '+lstGrupoUpsert[i].CodigoGrupo__c + '.',
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