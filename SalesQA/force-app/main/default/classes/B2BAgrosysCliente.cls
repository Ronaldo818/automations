public with sharing class B2BAgrosysCliente extends B2BAgrosys {

    public class FiltroRequisicao {
        public Boolean cliCommerce;
        public Boolean indUnderline;
        public Integer empresa;
        public Integer unidade;
        public Integer codCliente;
        public String nomeFantasia;
        public String razaoSocial;
        public String cnpj;
        public String cpf;
    }

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer cli_codigo;
        public String cli_fantasi;
        public String cli_razsoc;
        public String cli_cnpj;
        public String cli_cpf;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public List<TpEnvio> tp_envio;
    }

    public List<B2BAgrosysRetorno.TpCliente> getCliente(FiltroRequisicao filtro) {
        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'GET-CLIENTE';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.tp_envio = new list <TpEnvio>();

        TpEnvio tpEnvio = new TpEnvio();
        tpEnvio.cli_ecommerce = filtro.cliCommerce != null ? filtro.cliCommerce : true;
        tpEnvio.ind_underline = filtro.indUnderline != null ? filtro.indUnderline : true;
        tpEnvio.emp_empresa = filtro.empresa != null ? filtro.empresa : 1;
        tpEnvio.uni_unidade = filtro.unidade != null ? filtro.unidade : 0;
        tpEnvio.cli_codigo = filtro.codCliente != null ? filtro.codCliente : 0;
        tpEnvio.cli_fantasi = filtro.nomeFantasia != null ? filtro.nomeFantasia : '';
        tpEnvio.cli_razsoc = filtro.razaoSocial != null ? filtro.razaoSocial : '';
        tpEnvio.cli_cnpj = filtro.cnpj != null ? filtro.cnpj : '';
        tpEnvio.cli_cpf = filtro.cpf != null ? filtro.cpf : '';
        dadosReceb.tp_envio.add(tpEnvio);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serialize(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.method = 'POST';
        parametro.classMethod = 'getCliente';
        parametro.className = 'B2BAgrosysCliente';
        parametro.customerCode = String.valueOf(filtro.codCliente);
        parametro.integrationName = 'Clientes';
        
        B2BAgrosysRetorno retorno = String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
        if( retorno != null
            && !retorno.hasError
            && retorno.tp_retorno != null
            && !retorno.tp_retorno.isEmpty()
            && retorno.tp_retorno[0]?.msg_codigo == 1) {
            return retorno.tp_retorno[0].msg_retorno[0].tp_cliente;
        } else {
            return null;
        }
    }
}