/**
 * @author              : Allefy Joaquim
 * @group               : Sottama Tecnologia
 * @last modified on    : 2023-10-20
 * @last modified by    : Allefy Joaquim
 * 
**/

public without sharing virtual class B2BAgrosys {
    public Request dados;

    public class Request {
        public Requisicao request;
    }

    public class Requisicao {
        public List<DadosParam> vdadosParam;
        public List<DadosReceb> vdadosReceb;
    }

    public class DadosParam {
        public String tip_operacao;
    }

    public virtual class DadosReceb {
        // Os atributos abaixo serão definidos nas classes que implementam esta classe virtual.
        // public List<TpEnvio> tp_envio;
        // public DS ds;
    }

    public virtual class TpEnvio {
        public Integer emp_empresa;
        public Integer uni_unidade;
        public Boolean ind_underline;
        public Boolean cli_ecommerce;
    }
    
    public class DS {
        public List<TpPedido> tp_pedido;
    }

    public class TpPedido {
        public Boolean ind_underline;
        public Integer emp_empresa;
        public Integer uni_unidade;
        public String ped_codigo;
        public String ped_palm;
        public Integer cli_codigo;
        public Date ped_emissao;
        public Date ped_entrega;
        public String cpg_condpag;
        public Decimal ped_cashback;
        public String ped_observ;
        public String ped_obsnf;
        public Boolean ped_encerrado;
        public String ped_id;
        public Integer ped_operacao;
        public List<TpPedidoitem> tp_pedidoitem;
        public List<TpVencparc> tp_vencparc;
    }

    public class TpPedidoitem {
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

    public class TpVencparc {
        public Integer emp_empresa; 
        public Integer uni_unidade;
        public String tipo;
        public Integer nrparcela;
        public Date dtvenc;
        public Integer ped_codigo;
        public Integer nrdoc;
        public Decimal valor;
        public String parcent;
        public Decimal nrautoriz;
        public String tppag;
        public Integer bandeira;
        public String cnpjcartao;
        public Integer seq;
        public String chvpix;
    }

    public class TpCliente {
        public Integer emp_empresa;
        public Integer age_agencia01;
        public Integer age_agencia02;
        public Integer age_agencia03;
        public Integer age_agencia04;
        public Integer age_agencia05;
        public Integer age_agencia06;
        public Integer age_agencia07;
        public Integer age_agencia08;
        public Integer age_agencia09;
        public Integer age_agencia10;
        public String anacomerc;
        public String anafinanc;
        public String analogist;
        public String aniverfundac;
        public Integer bai_codent;
        public Integer bai_codigo;
        public String  bai_noment;
        public String  bairro;
        public Integer ban_banref01;
        public Integer ban_banref02;
        public Integer ban_banref03;
        public Integer ban_banref04;
        public Integer ban_banref05;
        public Integer ban_banref06;
        public Integer ban_banref07;
        public Integer ban_banref08;
        public Integer ban_banref09;
        public Integer ban_banref10;
        public Integer ccorrente01;
        public Integer ccorrente02;
        public Integer ccorrente03;
        public Integer ccorrente04;
        public Integer ccorrente05;
        public Integer ccorrente06;
        public Integer ccorrente07;
        public Integer ccorrente08;
        public Integer ccorrente09;
        public Integer ccorrente10;
        public String cep;
        public Integer cid_cidade;
        public Integer cid_cident;
        public Integer cli_catega;
        public String  cli_cep;
        public String  cli_cepent;
        public Integer cli_codmat;
        public Boolean cli_diaent01;
        public Boolean cli_diaent02;
        public Boolean cli_diaent03;
        public Boolean cli_diaent04;
        public Boolean cli_diaent05;
        public Boolean cli_diaent06;
        public Boolean cli_diaent07;
        public String cli_endent;
        public Integer cli_nrmaxdias;
        public String cli_nument;
        public Boolean cli_paletdesc;
        public Integer cli_subcat;
        public String cnpj;
        public Integer codcliori;
        public Integer codigo;
        public Integer codpalm;
        public String con_cargo;
        public String con_cpf;
        public String con_endereco;
        public String con_fone;
        public String con_identidade;
        public String con_nome;
        public Boolean con_sexo;
        public String cpf;
        public String cpg_condpag;
        public String ctt_cargo;
        public String ctt_cpf;
        public String ctt_endereco;
        public String ctt_fone1;
        public String ctt_identidade;
        public String ctt_nome;
        public Boolean ctt_sexo;
        public String digcc01;
        public String digcc02;
        public String digcc03;
        public String digcc04;
        public String digcc05;
        public String digcc06;
        public String digcc07;
        public String digcc08;
        public String digcc09;
        public String digcc10;
        public String digitacao;
        public String dtanacomerc;
        public String dtanafinanc;
        public String dtanalogist;
        public String email_contato;
        public String email_danfe;
        public String email_xml;
        public String endereco;
        public String fone;
        public String fonebco01;
        public String fonebco02;
        public String fonebco03;
        public String fonebco04;
        public String fonebco05;
        public String fonebco06;
        public String fonebco07;
        public String fonebco08;
        public String fonebco09;
        public String fonebco10;
        public Integer fxp_tpcli;
        public Decimal hlc_vllimlib;
        public String inscest;
        public Decimal limcresol;
        public String nome_fantasia;
        public String nom_cidade;
        public String estado;
        public String complem;
        public String numero;
        public String obs_anacomerc01;
        public String obs_anacomerc02;
        public String obs_anacomerc03;
        public String obs_anafinanc01;
        public String obs_anafinanc02;
        public String obs_anafinanc03;
        public String obs_analogist01;
        public String obs_analogist02;
        public String obs_analogist03;
        public String obs_contato;
        public String observ01;
        public String observ02;
        public String observ03;
        public Integer ocortabpre;
        public Integer palmcod;
        public String pcd_docprop01;
        public String pcd_docprop02;
        public String pcd_docprop03;
        public String pcd_docprop04;
        public String pcd_docprop05;
        public String pcd_docprop06;
        public String pcd_docprop07;
        public String pcd_docprop08;
        public String pcd_docprop09;
        public String pcd_docprop10;
        public String pcd_dtinclu;
        public String pcd_proprietarios01;
        public String pcd_proprietarios02;
        public String pcd_proprietarios03;
        public String pcd_proprietarios04;
        public String pcd_proprietarios05;
        public String pcd_proprietarios06;
        public String pcd_proprietarios07;
        public String pcd_proprietarios08;
        public String pcd_proprietarios09;
        public String pcd_proprietarios10;
        public String pontoref;
        public Integer por_codigo;
        public Integer pre_tiporeg;
        public Boolean predprop;
        public Integer ram_codigo;
        public String razsoc;
        public String refcomerc01;
        public String refcomerc02;
        public String refcomerc03;
        public String refcomerc04;
        public String refcomerc05;
        public String refcomerc06;
        public String refcomerc07;
        public String refcomerc08;
        public String refcomerc09;
        public String refcomerc10;
        public String rot_codigo;
        public Integer rvd_codigo;
        public Integer situac;
        public String stlib;
        public String tcl_codigo;
        public String telef01;
        public String telef02;
        public String telef03;
        public Boolean tpconta01;
        public Boolean tpconta02;
        public Boolean tpconta03;
        public Boolean tpconta04;
        public Boolean tpconta05;
        public Boolean tpconta06;
        public Boolean tpconta07;
        public Boolean tpconta08;
        public Boolean tpconta09;
        public Boolean tpconta10;
        public String tpd_tpdocmto;
        public Integer tpempr;
        public Integer tpp_tpportad;
        public String usu_anacomerc;
        public String usu_anafinanc;
        public String usu_analogist;
        public String usu_usuario;
        public Integer vis_codigo;
        public Boolean inscestise;
        public String rowid;
        public Boolean ind_underline;

        public TpCliente() {
            emp_empresa = 1;
            age_agencia01 = 0;
            age_agencia02 = 0;
            age_agencia03 = 0;
            age_agencia04 = 0;
            age_agencia05 = 0;
            age_agencia06 = 0;
            age_agencia07 = 0;
            age_agencia08 = 0;
            age_agencia09 = 0;
            age_agencia10 = 0;
            anacomerc = '';
            anafinanc = '';
            analogist = '';
            aniverfundac = '';
            bai_codent = 0;
            bai_codigo = 0;
            bai_noment = '';
            bairro = '';
            ban_banref01 = 0;
            ban_banref02 = 0;
            ban_banref03 = 0;
            ban_banref04 = 0;
            ban_banref05 = 0;
            ban_banref06 = 0;
            ban_banref07 = 0;
            ban_banref08 = 0;
            ban_banref09 = 0;
            ban_banref10 = 0;
            ccorrente01 = 0;
            ccorrente02 = 0;
            ccorrente03 = 0;
            ccorrente04 = 0;
            ccorrente05 = 0;
            ccorrente06 = 0;
            ccorrente07 = 0;
            ccorrente08 = 0;
            ccorrente09 = 0;
            ccorrente10 = 0;
            cep = '';
            cid_cidade = 0;
            cid_cident = 0;
            cli_catega = 0;
            cli_cep = '';
            cli_cepent = '';
            cli_codmat = 0;
            cli_diaent01 = false;
            cli_diaent02 = false;
            cli_diaent03 = false;
            cli_diaent04 = false;
            cli_diaent05 = false;
            cli_diaent06 = false;
            cli_diaent07 = false;
            cli_endent = '';
            cli_nrmaxdias = 0;
            cli_nument = '';
            cli_paletdesc = false;
            cli_subcat = 0;
            cnpj = '';
            codcliori = 0;
            codigo = 0;
            codpalm = 0;
            con_cargo = '';
            con_cpf = '';
            con_endereco = '';
            con_fone = '';
            con_identidade = '';
            con_nome = '';
            con_sexo = false;
            cpf = '';
            cpg_condpag = '';
            ctt_cargo = '';
            ctt_cpf = '';
            ctt_endereco = '';
            ctt_fone1 = '';
            ctt_identidade = '';
            ctt_nome = '';
            ctt_sexo = false;
            digcc01 = '';
            digcc02 = '';
            digcc03 = '';
            digcc04 = '';
            digcc05 = '';
            digcc06 = '';
            digcc07 = '';
            digcc08 = '';
            digcc09 = '';
            digcc10 = '';
            digitacao = '';
            dtanacomerc = '';
            dtanafinanc = '';
            dtanalogist = '';
            email_contato = '';
            email_danfe = '';
            email_xml = '';
            endereco = '';
            fone = '';
            fonebco01 = '';
            fonebco02 = '';
            fonebco03 = '';
            fonebco04 = '';
            fonebco05 = '';
            fonebco06 = '';
            fonebco07 = '';
            fonebco08 = '';
            fonebco09 = '';
            fonebco10 = '';
            fxp_tpcli = 0;
            hlc_vllimlib = 0;
            inscest = '';
            limcresol = 0;
            nome_fantasia = '';
            nom_cidade = '';
            estado = '';
            complem = '';
            numero = '';
            obs_anacomerc01 = '';
            obs_anacomerc02 = '';
            obs_anacomerc03 = '';
            obs_anafinanc01 = '';
            obs_anafinanc02 = '';
            obs_anafinanc03 = '';
            obs_analogist01 = '';
            obs_analogist02 = '';
            obs_analogist03 = '';
            obs_contato = '';
            observ01 = '';
            observ02 = '';
            observ03 = '';
            ocortabpre = 0;
            palmcod = 0;
            pcd_docprop01 = '';
            pcd_docprop02 = '';
            pcd_docprop03 = '';
            pcd_docprop04 = '';
            pcd_docprop05 = '';
            pcd_docprop06 = '';
            pcd_docprop07 = '';
            pcd_docprop08 = '';
            pcd_docprop09 = '';
            pcd_docprop10 = '';
            pcd_dtinclu = '';
            pcd_proprietarios01 = '';
            pcd_proprietarios02 = '';
            pcd_proprietarios03 = '';
            pcd_proprietarios04 = '';
            pcd_proprietarios05 = '';
            pcd_proprietarios06 = '';
            pcd_proprietarios07 = '';
            pcd_proprietarios08 = '';
            pcd_proprietarios09 = '';
            pcd_proprietarios10 = '';
            pontoref = '';
            por_codigo = 0;
            pre_tiporeg = 0;
            predprop = false;
            ram_codigo = 0;
            razsoc = '';
            refcomerc01 = '';
            refcomerc02 = '';
            refcomerc03 = '';
            refcomerc04 = '';
            refcomerc05 = '';
            refcomerc06 = '';
            refcomerc07 = '';
            refcomerc08 = '';
            refcomerc09 = '';
            refcomerc10 = '';
            rot_codigo = '';
            rvd_codigo = 0;
            situac = 0;
            stlib = '';
            tcl_codigo = '';
            telef01 = '';
            telef02 = '';
            telef03 = '';
            tpconta01 = false;
            tpconta02 = false;
            tpconta03 = false;
            tpconta04 = false;
            tpconta05 = false;
            tpconta06 = false;
            tpconta07 = false;
            tpconta08 = false;
            tpconta09 = false;
            tpconta10 = false;
            tpd_tpdocmto = '';
            tpempr = 0;
            tpp_tpportad = 0;
            usu_anacomerc = '';
            usu_anafinanc = '';
            usu_analogist = '';
            usu_usuario = '';
            vis_codigo = 0;
            inscestise = false;
            rowid = '';
            ind_underline = true;
        }
    }

    public class HttpParameters {
        public String body;
        public String method;
        public String integrationName;
        public String customerCode;
        public String classMethod;
        public String className;
        public String orderId;
        public String cartId;
        public String accountId;
        public IntegrationLog__c log;
    }

    // public class OccurredError {
    //     public Boolean hasError = false;
    //     public String errorMesage;
    // }

    public class RetornoWrapper {
        public Boolean hasError = false;
        public String errorMessage;
        public String logId;
        public String json;    
    }
	
    public static B2BAgrosysRetorno postRequest(HttpParameters parameters) {
        WebService__mdt service = WebService__mdt.getInstance('Agrosys');
        Http http = new Http();
        HttpRequest request = new HttpRequest();
        HttpResponse response = new HttpResponse();

        try{
            request.setHeader('Content-Type', 'application/json');
            request.setHeader('Accept', '*/*');
            request.setMethod(parameters.method);
            request.setBody(parameters.body);
            request.setEndpoint(service.BaseUrl__c);
            request.setTimeout(100000);
            response = http.send(request);
            
            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c = parameters.customerCode != null ? parameters.integrationName+' - '+parameters.customerCode : parameters.integrationName;
			log.ApexClass__c = parameters.className;
            log.ApexMethod__c = parameters.classMethod;
            log.OrderId__c = String.isNotBlank(parameters.orderId) ? parameters.orderId : null;
            log.AccountId__c = String.isNotBlank(parameters.accountId) ? parameters.accountId : null;
            log.CartId__c = String.isNotBlank(parameters.cartId) ? parameters.cartId : null;
			log.EndPoint__c = request.getEndpoint();
			log.IntegrationMethod__c = request.getMethod();
			log.IntegrationStatusCode__c = response.getStatusCode()+' : '+response.getStatus();
			log.RequestBody01__c = request.getBody();
			log.OperationStatus__c = response.getStatusCode() != 200 && response.getStatusCode() != 201 ? 'Error' : 'Success';

            if(String.isNotBlank(parameters.cartId)) {
                log.CartId__c = parameters.cartId;
            }
            if(String.isNotBlank(parameters.orderId)) {
                log.OrderId__c = parameters.orderId;
            }
            if(String.isNotBlank(parameters.AccountId)) {
                log.AccountId__c = parameters.accountId;
            }
			
            if(response.getBody().length() < 131000) {
                log.responseBody01__c =  response.getBody();
            }
            else if(response.getBody().length() < 131000 * 2) {
                log.responseBody01__c =  response.getBody().substring(0, 131072);
                log.responseBody02__c =  response.getBody().substring(131072, response.getBody().length());
            }
            else if(response.getBody().length() < 131000 * 3) {
                log.responseBody01__c = response.getBody().substring(0, 131072);
                log.responseBody02__c = response.getBody().substring(131072, 131072*2);
                log.responseBody03__c = response.getBody().substring(262143, response.getBody().length());
            }
            else {
                log.responseBody01__c = response.getBody().substring(0, 131072);
                log.responseBody02__c = response.getBody().substring(131072, 131072*2);
                log.responseBody03__c = response.getBody().substring(262143, 262143+131072);
            }
            insert log;
            parameters.log = log;

            if(response.getStatusCode() != 200 && response.getStatusCode() != 201) {
                B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
                retorno.hasError = true;
                retorno.logId = log.Id;
                retorno.errorMessage = 'Ocorreu um erro, por favor tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema.';
                return retorno;
            }
            if(!response.getBody().startsWith('{')){
                B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
                retorno.hasError = true;
                retorno.tp_retorno = null;
                retorno.logId = log.Id;
                return null;
            }
             else {
                B2BAgrosysRetorno retorno = B2BAgrosysRetorno.parse(response.getBody(), parameters);
                retorno.logId = retorno.logId != null ? retorno.logId : log.Id;
                return retorno;
            }
        }
        catch(Exception e){
            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c = parameters.customerCode != null ? parameters.integrationName+' - '+parameters.customerCode : parameters.integrationName;
			log.ApexClass__c = parameters.className;
            log.ApexMethod__c = parameters.classMethod;
            log.OrderId__c = String.isNotBlank(parameters.orderId) ? parameters.orderId : null;
            log.AccountId__c = String.isNotBlank(parameters.accountId) ? parameters.accountId : null;
			log.EndPoint__c = request.getEndpoint();
			log.IntegrationMethod__c = request.getMethod();
			log.IntegrationStatusCode__c = response.getStatusCode()+' : '+response.getStatus();
			log.RequestBody01__c = request.getBody();
			log.OperationStatus__c = response.getStatusCode() != 200 || response.getStatusCode() != 201 ? 'Error' : 'Success';
            log.MessageError__c = 'Type: ' + e.getTypeName() + '\n\nCause: ' + e.getCause() + '\n\nLine Number: ' + e.getLineNumber() + '\n\nStack Trace: ' + e.getStackTraceString() + '\n\nMessage: ' + e.getMessage();
			
            if(response.getBody().length() < 131000) {
                log.responseBody01__c =  response.getBody();
            }else if(response.getBody().length() < 131000 * 2) {
                log.responseBody01__c =  response.getBody().substring(0, 131072);
                log.responseBody02__c =  response.getBody().substring(131072, response.getBody().length());
            } else if(response.getBody().length() < 131000 * 3) {
                log.responseBody01__c = response.getBody().substring(0, 131072);
                log.responseBody02__c = response.getBody().substring(131072, 131072*2);
                log.responseBody03__c = response.getBody().substring(262143, response.getBody().length());
            } else {
                log.responseBody01__c = response.getBody().substring(0, 131072);
                log.responseBody02__c = response.getBody().substring(131072, 131072*2);
                log.responseBody03__c = response.getBody().substring(262143, 262143+131072);
            }
            insert log;

            System.debug('ERROR B2BAgrosys => \n' + log.MessageError__c);

            B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
            retorno.hasError = true;
            retorno.errorMessage = 'Ocorreu um erro. Por favor, tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema.';
            retorno.logId = log.Id;
            return retorno;
        }
    }


    public static String getErrorAgrosys(Integer codigo) {
        try {
            List<MessageErrorAgrosys__mdt> message = [    
                SELECT MessageCode__c, MessageDescription__c
                FROM MessageErrorAgrosys__mdt
                WHERE MessageCode__c = :codigo
            ];
            if(!message.isEmpty()) {
                return message[0].MessageDescription__c;
            }
            else {
                MessageErrorAgrosys__mdt defaultErrorMessage = [
                    SELECT MessageCode__c, MessageDescription__c
                    FROM MessageErrorAgrosys__mdt
                    WHERE MessageCode__c = 999
                ];
                return defaultErrorMessage.MessageDescription__c; 
            }
        }
        catch(Exception ex) {
            return ex.getMessage();
        }
    }

}