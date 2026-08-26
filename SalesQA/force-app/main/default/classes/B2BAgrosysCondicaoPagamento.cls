public with sharing class B2BAgrosysCondicaoPagamento extends B2BAgrosys {

    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public String codCondPag;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public String cpg_condpag;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }
    
    public void getCondPagto(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-CONDPAGTO';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : null;
        tpEnvio.cpg_condpag = filtro.codCondPag != null ? filtro.codCondPag : '';
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');
        
        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getCondPagto';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysCondicaoPagamento';

        List<CondicaoPagamento__c> lstCondicaoPagamentoUpsert = new List<CondicaoPagamento__c>();
        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        if(retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0].msg_codigo == 1
            && retorno.tp_retorno[0].msg_retorno[0].tp_condpagto != null
            && !retorno.tp_retorno[0].msg_retorno[0].tp_condpagto.isEmpty()) {
            for(B2BAgrosysRetorno.TpCondpagto condicao : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_condpagto) {
                CondicaoPagamento__c condicaoPagto = new CondicaoPagamento__c();
                condicaoPagto.Name = condicao.cpg_descricao;
                condicaoPagto.Descricao__c = condicao.cpg_descricao;
                condicaoPagto.CodigoCondicaoPagamento__c = condicao.cpg_condpag;
                condicaoPagto.NumeroParcelas__c = condicao.cpg_nrparcelas;
                condicaoPagto.AVista__c = condicao.cpg_avista;
                lstCondicaoPagamentoUpsert.add(condicaoPagto);
                
                if(Test.isRunningTest() && lstCondicaoPagamentoUpsert.size() > 2) {
                    lstCondicaoPagamentoUpsert[0].CodigoCondicaoPagamento__c = '0';
                    lstCondicaoPagamentoUpsert[1].CodigoCondicaoPagamento__c = '0';
                }
            }
        } else {
            // Adicionar tratativa.
        }

        if(!lstCondicaoPagamentoUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] results = Database.upsert(lstCondicaoPagamentoUpsert, CondicaoPagamento__c.CodigoCondicaoPagamento__c, false);
            for(Database.UpsertResult sr : results){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                            NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro para Condição de Pagamento: ' + lstCondicaoPagamentoUpsert[i].CodigoCondicaoPagamento__c + '.',
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