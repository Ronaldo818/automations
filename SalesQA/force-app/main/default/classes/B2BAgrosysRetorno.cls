/*
 * @author              : Allefy Joaquim
 * @group               : Sottama Tecnologia
 * @last modified on    : 2023-10-17
 * @last modified by    : Allefy Joaquim
 */

public with sharing class B2BAgrosysRetorno {
    
    public List<TpRetorno> tp_retorno;
    public Boolean hasError = false;
    public String errorMessage;
    public String logId;

    public static B2BAgrosysRetorno parse(String json, B2BAgrosys.HttpParameters parameters) {
        try {
            B2BAgrosysRetorno retorno = (B2BAgrosysRetorno) System.JSON.deserialize(json, B2BAgrosysRetorno.class);
            retorno.hasError = false;
            return retorno;
        } catch(Exception e) {
            IntegrationLog__c log = new IntegrationLog__c();

            if(parameters.log != null) {
                log = parameters.log;
                log.NomeLog__c = 'ERRO ' + log.NomeLog__c;
                log.ApexClass__c = 'B2BAgrosysRetorno';
                log.ApexMethod__c = 'parse';
			    log.OperationStatus__c = 'Error';
                log.MessageError__c = 'Type: ' + e.getTypeName() + '\n\nCause: ' + e.getCause() + '\n\nLine Number: ' + e.getLineNumber() + '\n\nStack Trace: ' + e.getStackTraceString() + '\n\nMessage: ' + e.getMessage();
            } else {
                log.NomeLog__c = parameters.customerCode != null ? 'ERRO com a integração Agrosys - ' + parameters.customerCode : 'ERRO com a integração Agrosys';
                log.OrderId__c = String.isNotBlank(parameters.orderId) ? parameters.orderId : null;
                log.AccountId__c = String.isNotBlank(parameters.accountId) ? parameters.accountId : null;
                log.CartId__c = String.isNotBlank(parameters.cartId) ? parameters.cartId : null;
                log.ApexClass__c = 'B2BAgrosysRetorno';
                log.ApexMethod__c = 'parse';
			    log.OperationStatus__c = 'Error';
                log.MessageError__c = 'Type: ' + e.getTypeName() + '\n\nCause: ' + e.getCause() + '\n\nLine Number: ' + e.getLineNumber() + '\n\nStack Trace: ' + e.getStackTraceString() + '\n\nMessage: ' + e.getMessage();
            }
            upsert log;

            System.debug('ERROR B2BAgrosysRetorno => \n' + log.MessageError__c);
            
            B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
            retorno.hasError = true;
            retorno.errorMessage = 'Ocorreu um erro. Por favor, tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema.';
            retorno.logId = log.Id;
            return retorno;
        }
    }

    public class TpRetorno {
    	public Integer msg_codigo;
        public String msg_descricao;
        public List<MensagemRetorno> msg_retorno;
    }

    public class MensagemRetorno {
        public List<TpCidade> tp_cidade; // Cidade
        public List<tpRamo> tp_rm; // Ramo de Atividade
        public List<TpCondpagto> tp_condpagto; // Condição de Pagamento
        public List<TpFamilia> tp_familia; // Família
        public List<TpGrpitem> tp_grpitem; // // Grupo do Item
        public List<tpUnidade> tp_un; // Unidade de Medida
        public List<TpLinha> tp_linha; // Linha
        public List<TpOrigem> tp_origem; // Origem
        public List<TabPreco> tp_tabpreco; // Tabela de Preço
        public List<TpPedido> tp_pedido;
        public List<TpPedidoItem> tp_peditem;
        public List<TpCliente> tp_cliente; // Cliente
        public List<TpProduto> tp_produto; // Produto
        public List<Titulo> tp_titulo; // Títulos
        public List<TpEstoque> tp_estoque; // Estoque
        public List<TpEndereco> tp_endereco; // Endereço
        public DS ds;
        public List<TpExtverba> tp_extverba; // Cashback
        public Integer status;
        public String msg;
        public Obj01 obj; // campo utilizado no retorno da criação de um pedido
        public List<TpCliNeg> tp_clineg; // Negocioação do cliente.
    }

    public class TpCliente {
        public Integer cli_codigo;
        public Integer emp_empresa;
        public Integer uni_unidade;
        public String cli_fantasi;
        public String cli_razsoc;
        public Boolean cli_fisjur;
        public String cli_cnpj;
        public String cli_cpf;
        public Boolean cli_aplicast;
        public String cli_email;
        public String cli_foncom;
        public String cli_fonfin;
        public String cli_dtativ;
        public String cli_dtinativ;
        public Decimal cli_overprice;
        public Decimal cli_vlcliverba;
        public Decimal cli_limcre;
        public String cpg_condpag; // tabela Condiçao de Pagamento
        public String cpg_descri; // Descrição da Condição de Pagamento
        public Integer fxp_tpcli;
        public Integer pre_tiporeg;
        public Integer pit_tabpre;
        public Boolean cli_ecommerce;
        public Decimal ped_descon;
        public Decimal cli_descfin;
    }

    public class TpCondpagto {
        public String cpg_condpag;
		public String cpg_descricao;
		public Boolean cpg_avista;
		public Integer cpg_nrparcelas;
        public Integer emp_empresa;
    }

    public class TpProduto {
        public Integer pro_codigo;
        public String pro_descricao;
        public Integer emp_empresa;
        public Integer uni_unidade;
        public String med_codigo;
        public Decimal pro_psunitario;
        public Boolean pro_situa;
        public Integer fam_codigo; // tabela Família
        public Integer gru_codigo; // tabela Grupo do Item
        public Integer lin_codigo; // tabela Linha
        public Integer ori_codigo; // tabela Origem
        public Integer emb_peso;
        public Integer emb_altura;
        public Integer emb_largura;
        public Integer emb_comprimento;
        public String pro_uniped;
        public String pro_unifat;
    }

    public class TpFamilia {
        public Integer fam_codigo;
        public String fam_descricao;
    }

    public class TpGrpitem {
        public Integer gru_codigo;
        public String gru_descricao;
    }

    public class TpLinha {
        public Integer lin_codigo;
        public String lin_descricao;
    }

    public class TpOrigem {
        public Integer ori_codigo;
        public String ori_descricao;
    }

    public class Titulo {
        public Integer emp_empresa;
        public Integer fil_filial;
        public Integer uni_unidade;
        public Integer cli_codigo;
        public String tit_numuni;
        public Integer tit_nrparcela;
        public Date tit_dtemissao;
        public Date tit_dtvencto;
        public Date tit_dtpagto;
        public String nfs_serie;
        public Integer nfs_numero;
        public Decimal tit_vldocmto;
        public Decimal tit_vlpagto;
        public Decimal tit_vlpagtoparc;
        public Decimal tit_vlsaldodev;
    }

    public class TpCidade {
        public Integer cid_cidade;
        public String cid_nome;
        public Integer cid_ibge;
        public String est_estado;
        public Integer pai_pais;
    }

    public class tpRamo {
        public Integer ram_codigo;
        public String ram_nome;
        public String ram_sigla;
    }

    public class tpUnidade {
        public String med_codigo;
        public String med_descricao;
        public Boolean med_tipo;
    }

    public class TabPreco {
        public Integer emp_empresa;
        public Integer uni_unidade;
        public Integer cli_codigo;
        public Integer pre_tiporeg;
        public Integer pit_tabpre;
        public Integer pro_codigo;
        public Date pre_dtatual;
        public Decimal pre_tabela;
        public Decimal pre_vlrst;
        public Decimal pre_vfcpst;
    }

    public class TpEstoque {
        public Integer emp_empresa;
        public Integer uni_unidade;
        public Integer pro_codigo;
        public Date est_dtdisp;
        public Decimal est_camara;
    }

    public class DS {
        public List<TpPedido> tp_pedido;
    }

    public class TpPedido {
        public Integer emp_empresa;
        public Integer uni_unidade;
        public String ped_codigo;
        public String ped_palm;
        public Integer cli_codigo;
        public Date ped_dtemiss;
        public Date ped_dtentrega;
        public String ped_status;
        public Date ped_emissao;
        public Date ped_entrega;
        public String cpg_condpag;
        public Decimal ped_cashback;
        public String ped_observ;
        public String ped_obsnf;
        public Boolean ped_encerrado;
        public Boolean ind_underline;
        public List<TpPedidoItem> tp_pedidoitem;
    }

    public class TpPedidoItem {
        public Integer emp_empresa;
        public Integer uni_unidade;
        public String ped_codigo;
        public Integer pro_codigo;
        public String med_codigo;
        public Decimal pit_qtpedi;
        public Decimal pit_pspedi;
        public Decimal pit_vlpreco;
        public Integer pit_seq;
        public String pit_observa;
        public Integer pre_tiporeg;
        public Integer pit_tabpre;
        public Integer pit_status;
        public String pit_ordcomp;
        public Integer pit_seqordcomp;
        public Decimal pit_vlprecoaux;
        public Decimal pit_vlrst;
        public Decimal pit_vfcpst;
        public Decimal pit_vldesc;
    }

    public class TpEndereco {
        public Integer emp_empresa;
        public Integer uni_unidade;
        public Integer cli_codigo;
        public Integer ced_tpender;
        public Integer ced_seq;
        public String ced_endereco;
        public String ced_numero;
        public String ced_cep;
        public String cli_bairro;
        public Integer cid_cidade;
        public String end_latitude;
        public String end_longitude;
        public String end_padrao;
        public String end_descricao;
        public String ced_observ;
    }

    public class TpExtverba {
        public String cli_codigo;
        public Date hvr_dtmov;
        public Decimal hvr_tpverba;
        public Boolean hvr_tpmov;
        public String hvr_tpdoc;
        public String hvr_docto;
        public Decimal hvr_valor;
        public Decimal ite_codigo;
        public Decimal hvr_saldo;
        public Integer hvr_seq;
    }

    public class Obj01 {
        public Decimal ped_codigo;
        public MensIntegra mensintegra;
        public List<RetPed> retped;
        public Integer vis_codigo;
    }

    public class RetPed {
        public Integer pit_seq;        // Sequência do item no pedido.
        public Decimal pit_qtdpedi;    // Quantidade desejada do item.
        public Decimal pit_estatua;
        public String ped_codigo;     // Número do pedido. 
        public String ped_erro;        // Causa do erro.
        public String sta_erro;        // Indica se o erro foi no pedido (cód 98), ou no item do pedido (cód 99).
        public String cod_erro;        // Código do erro.    
    }

    public class MensIntegra {
        public Integer status;
        public String msg;
        public String vlTEFF;
        public Obj02 obj;
    }
    
    public class Obj02 {
        public Integer vis_codigo;
    }

    public class TpCliNeg {
        public Boolean temnegociacao;
    }
}