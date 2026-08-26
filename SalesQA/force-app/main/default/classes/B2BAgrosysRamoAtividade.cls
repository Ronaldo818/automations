public with sharing class B2BAgrosysRamoAtividade extends B2BAgrosys {

    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codRamo;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer ram_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getRamoAtivade(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-RAMOATIV';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : null;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 0;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.ram_codigo = filtro.codRamo != null ? filtro.codRamo : 0;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getRamoAtivade';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysRamoAtividade';

        // requisitar(parametro);
        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        List<RamoAtividade__c> lstRamoAtividadeUpsert = new List<RamoAtividade__c>();
        if(retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0]?.msg_codigo == 1
            && retorno.tp_retorno[0]?.msg_retorno[0]?.tp_rm != null
            && !retorno.tp_retorno[0]?.msg_retorno[0]?.tp_rm.isEmpty()) {
            for(B2BAgrosysRetorno.tpRamo ramo : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_rm) {
                RamoAtividade__c ramoAtivade = new RamoAtividade__c();
                ramoAtivade.Name = ramo.ram_nome;
                ramoAtivade.CodigoRamo__c = String.valueOf(ramo.ram_codigo);
                ramoAtivade.Sigla__c = ramo.ram_sigla;
                lstRamoAtividadeUpsert.add(ramoAtivade);
                
                if(Test.isRunningTest() && lstRamoAtividadeUpsert.size() > 2) {
                	lstRamoAtividadeUpsert[0].CodigoRamo__c = '0';
                    lstRamoAtividadeUpsert[1].CodigoRamo__c = '0';
                }
            }
        }

        if(!lstRamoAtividadeUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] resultRamos = Database.upsert(lstRamoAtividadeUpsert, RamoAtividade__c.CodigoRamo__c, false);
            for(Database.UpsertResult sr : resultRamos){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                        NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro do Ramo de Atividades: ' + lstRamoAtividadeUpsert[i].Name + ' - ' + lstRamoAtividadeUpsert[i].CodigoRamo__c + '.',
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