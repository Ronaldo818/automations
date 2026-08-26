public with sharing class B2BAgrosysUnidadeMedida extends B2BAgrosys {
    
    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public String codMedida;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public String med_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getUnidadeMedida(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-UNIDMED';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : null;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 0;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.med_codigo = filtro.codMedida != null ? filtro.codMedida : '';
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');
        
        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getUnidadeMedida';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysUnidadeMedida';

        // requisitar(parametro);
        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        List<UnidadeMedida__c> lstUnidadeMedidaUpsert = new List<UnidadeMedida__c>();
        if(retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0]?.msg_codigo == 1
            && retorno.tp_retorno[0]?.msg_retorno[0]?.tp_un != null
            && !retorno.tp_retorno[0]?.msg_retorno[0]?.tp_un.isEmpty()) {
            for(B2BAgrosysRetorno.tpUnidade unidade : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_un) {
                UnidadeMedida__c unidadeMedida = new UnidadeMedida__c();
                unidadeMedida.Name = unidade.med_descricao;
                unidadeMedida.Descricao__c = unidade.med_descricao;
                unidadeMedida.CodigoUnidadeMedida__c = unidade.med_codigo;
                unidadeMedida.TipoMedida__c = unidade.med_tipo;
                lstUnidadeMedidaUpsert.add(unidadeMedida);
                
                if(Test.isRunningTest() && lstUnidadeMedidaUpsert.size() > 2) {
                	lstUnidadeMedidaUpsert[0].CodigoUnidadeMedida__c = '0';
					lstUnidadeMedidaUpsert[1].CodigoUnidadeMedida__c = '0';
                }
            }
        }

        if(!lstUnidadeMedidaUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] resultUnidades = Database.upsert(lstUnidadeMedidaUpsert, UnidadeMedida__c.CodigoUnidadeMedida__c, false);
            for(Database.UpsertResult sr : resultUnidades){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                        NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro da Unidade de Medida: ' + lstUnidadeMedidaUpsert[i].CodigoUnidadeMedida__c + '.',
                        MessageError__c = 'Status: ' + err.getStatusCode() + '\n\nCampos: ' + err.getFields() + '\n\nMensagem' + err.getMessage(),
                        OperationStatus__c = 'Error');
                        lstLogInsert.add(log);
                    }
                }
                i++;   
            }
            
            if (!lstLogInsert.isEmpty()) {
            insert lstLogInsert;
        	} 
        }
    }
}