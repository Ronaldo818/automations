public without sharing class B2BAgrosysEstoqueProduto extends B2BAgrosys {

    public class FiltroRequisicao {
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codProduto;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer pro_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public B2BAgrosysRetorno getEstoque(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-ESTOQUE';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();
        
        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : null;
        tpEnvio.pro_codigo = filtro.codProduto != null ? filtro.codProduto : 0;

        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.classMethod = 'getEstoque';
        parametro.method = 'POST';
        parametro.className = 'B2BAgrosysEstoque';

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        if(retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0].msg_codigo == 1
            && retorno.tp_retorno[0].msg_retorno[0].tp_estoque != null
            && !retorno.tp_retorno[0].msg_retorno[0].tp_estoque.isEmpty()) {
            return retorno;
        } else {
            return null;
        }
    }
}