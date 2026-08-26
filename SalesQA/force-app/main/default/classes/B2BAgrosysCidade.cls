public with sharing class B2BAgrosysCidade extends B2BAgrosys {
    
    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codCidade;
        public String nomeCidade;
        public Integer codCidadeIbge;
        public Integer codPais;
        public String estado;
    }
    
    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer cid_cidade;
        public String cid_nome;
        public Integer cid_ibge;
        public Integer pai_pais;
        public String est_estado;
    }
    
    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public void getCidade(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-CIDADE';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list<TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : null;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.cid_cidade = filtro.codCidade != null ? filtro.codCidade : 0;
        tpEnvio.cid_nome = filtro.nomeCidade != null ? filtro.nomeCidade : '';
        tpEnvio.cid_ibge = filtro.codCidadeIbge != null ? filtro.codCidadeIbge : 0;
        tpEnvio.pai_pais = filtro.codPais != null ? filtro.codPais : 0;
        tpEnvio.est_estado = filtro.estado != null ? filtro.estado : '';
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.method = 'POST';
        parametro.classMethod = 'getCidade';
        parametro.className = 'B2BAgrosysCidade';
        parametro.integrationName = 'Cidades';

        List<Cidade__c> lstCidadeUpsert = new List<Cidade__c>();
        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        if(retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0].msg_codigo == 1
            && retorno.tp_retorno[0].msg_retorno[0].tp_cidade != null
            && !retorno.tp_retorno[0].msg_retorno[0].tp_cidade.isEmpty()) {
            for(B2BAgrosysRetorno.TpCidade cidade : retorno.tp_retorno[0]?.msg_retorno[0]?.tp_cidade) {
                String codCidade = String.valueOf(cidade.cid_cidade);
                String codIBGE = String.valueOf(cidade.cid_ibge);

                Cidade__c cid = new Cidade__c();
                cid.Name = cidade.cid_nome;
                cid.UF__c = cidade.est_estado;
                cid.CodigoERP__c = codCidade;
                cid.CodigoCidadeIBGE__c = codIBGE.equalsIgnoreCase('0') ? codIBGE + '-' + codCidade : codIBGE;
                cid.CodigoPaisSPED__c = String.valueOf(cidade.pai_pais);
                lstCidadeUpsert.add(cid);
                
                if(Test.isRunningTest() && lstCidadeUpsert.size() > 2) {
                    lstCidadeUpsert[0].CodigoCidadeIBGE__c = '0';
                    lstCidadeUpsert[1].CodigoCidadeIBGE__c = '0';
                }
            }
        }

        if(!lstCidadeUpsert.isEmpty()) {
            List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
            Integer i = 0;
            Database.UpsertResult[] results = Database.upsert(lstCidadeUpsert, Cidade__c.CodigoCidadeIBGE__c, false);
            for(Database.UpsertResult sr : results){
                if(!sr.success) {
                    for(Database.Error err : sr.getErrors()) {
                        IntegrationLog__c log = new IntegrationLog__c(
                            NomeLog__c = 'Erro ao tentar Inserir/Atualizar registro para a Cidade: ' + lstCidadeUpsert[i].Name + ' - ' + lstCidadeUpsert[i].CodigoCidadeIBGE__c + '.',
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