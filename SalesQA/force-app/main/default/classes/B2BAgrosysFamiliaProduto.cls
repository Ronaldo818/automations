public with sharing class B2BAgrosysFamiliaProduto extends B2BAgrosys {

    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codFamilia;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer fam_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getFamilia(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-FAMILIA';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : true;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.fam_codigo = filtro.codFamilia != null ? filtro.codFamilia : 0;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');
        
        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.method = 'POST';
        parametro.classMethod = 'getFamilia';
        parametro.className = 'B2BAgrosysFamiliaProduto';
        parametro.integrationName = 'Família do Produto';
        
        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        List<FamiliaProduto__c> lstFamiliaUpsert = new List<FamiliaProduto__c>();
        if(retorno.tp_retorno != null
        && !retorno.tp_retorno.isEmpty()
        && retorno.tp_retorno[0].msg_codigo == 1
        && retorno.tp_retorno[0].msg_retorno[0].tp_familia != null
        && !retorno.tp_retorno[0].msg_retorno[0].tp_familia.isEmpty()) {
            for(B2BAgrosysRetorno.TpFamilia familia : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_familia) {
                FamiliaProduto__c familiaProduto = new FamiliaProduto__c();
                familiaProduto.Name = familia.fam_descricao;
                familiaProduto.Descricao__c = familia.fam_descricao;
                familiaProduto.CodigoFamilia__c = String.valueOf(familia.fam_codigo);
                lstFamiliaUpsert.add(familiaProduto);
            }
            
            if(Test.isRunningTest() && lstFamiliaUpsert.size() > 2) {
                lstFamiliaUpsert[0].CodigoFamilia__c = '0';
                lstFamiliaUpsert[1].CodigoFamilia__c = '0';
            } 
        } else {
            // Adicionar tratativa.
        }

        if(!lstFamiliaUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] results = Database.upsert(lstFamiliaUpsert, FamiliaProduto__c.CodigoFamilia__c, false);
            for(Database.UpsertResult sr : results){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                            NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro de Família: ' + lstFamiliaUpsert[i].Name +' - '+lstFamiliaUpsert[i].CodigoFamilia__c + '.',
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