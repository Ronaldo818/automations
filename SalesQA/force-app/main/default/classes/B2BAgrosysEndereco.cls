public with sharing class B2BAgrosysEndereco extends B2BAgrosys {

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer cli_codigo;
        public Integer ced_tpender;
    }

    public List<B2BAgrosysRetorno.TpEndereco> getEnderecoCliente(Integer codigoCliente) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-ENDERECO';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();

        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.emp_empresa = 1;
        tpEnvio.uni_unidade = 0;
        tpEnvio.cli_codigo = codigoCliente;
        tpEnvio.ced_tpender = 0;
        tpEnvio.ind_underline = true;
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.method = 'POST';
        parametro.classMethod = 'getEnderecoCliente';
        parametro.className = 'B2BAgrosysEndereco';
        parametro.customerCode = String.valueOf(codigoCliente);
        parametro.integrationName = 'Endereços';

        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;        
        if(retorno.tp_retorno != null && !retorno.tp_retorno.isEmpty() && retorno.tp_retorno[0].msg_codigo == 1) {
            return retorno.tp_retorno[0].msg_retorno[0].tp_endereco;
        } else {
            return null;
        }
    }
}