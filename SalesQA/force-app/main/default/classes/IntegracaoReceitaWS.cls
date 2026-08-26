global without sharing class IntegracaoReceitaWS {

    static Map <String, String> estados = new Map<String, String>{ 'AC' => 'Acre', 'AL' => 'Alagoas', 'AP' => 'Amapá',
                                                             'AM' => 'Amazonas', 'BA' => 'Bahia', 'CE' => 'Ceará',
                                                             'ES' => 'Espírito Santo', 'GO' => 'Goiás', 
                                                             'MA' => 'Maranhão', 'MT' => 'Mato Grosso', 
                                                             'MS' => 'Mato Grosso do Sul', 'MG' => 'Minas Gerais',
                                                             'PA' => 'Pará', 'PB' => 'Paraíba', 'PR' => 'Paraná',
                                                             'PE' => 'Pernambuco', 'PI' => 'Piauí', 
                                                             'RJ' => 'Rio de Janeiro', 'RN' => 'Rio Grande do Norte',
                                                             'RS' => 'Rio Grande do Sul', 'RO' => 'Rondônia', 
                                                             'RR' => 'Roraima', 'SC' => 'Santa Catarina', 
                                                             'SP' => 'São Paulo', 'SE' => 'Sergipe', 
                                                             'TO' => 'Tocantins', 'DF' => 'Distrito Federal'};
   
     @AuraEnabled 
     public static Wrapper consultarDadosReceitaWS(String CNPJ) {
   
       CNPJ = CNPJ.replace( '/' , '' );
       CNPJ = CNPJ.replace( '-' , '' );
       CNPJ = CNPJ.replace( '.' , '' );
       CNPJ = CNPJ.replace( ' ' , '' );
   
       final String endpoint ='https://receitaws.com.br/v1/cnpj/'+ CNPJ;
       Http http = new Http();
       HttpRequest request = new HttpRequest();
   
       request.setEndpoint(endpoint);
       request.setMethod('GET');
       HttpResponse response = http.send(request);
   
       Wrapper wrapper = new Wrapper();
       if(response.getStatusCode() == 200) {
   
         DadosReceitaWS dados = (DadosReceitaWS)JSON.deserialize(response.getBody(), DadosReceitaWS.class);
   
         dados.cep = dados.cep.replace( '.' , '' );
   
         wrapper.dadosGeral = dados;
   
         wrapper.dadosGeral.uf = IntegracaoReceitaWS.estados.get(wrapper.dadosGeral.uf);
   
         String atividadePrim='<table class="slds-table slds-table_cell-buffer slds-table_bordered" aria-labelledby="element-with-table-label other-element-with-table-label">';
         atividadePrim +='<tr style="background-color:#eeeeee;" class="slds-line-height_reset">'+
                         '<th>Código</th><th style="align:left">Descrição</th></tr>';
         for(Atividade_principal atvPrim : dados.atividade_principal) {
             atividadePrim += '<tr> <td>' +  atvPrim.code + '</td><td> ' + atvPrim.text + '</td></tr>';
         }
         atividadePrim +='</table>';
         wrapper.dadosAtividadePrim = atividadePrim;
   
         String atividadesSecund='<table class="slds-table slds-table_cell-buffer slds-table_bordered" aria-labelledby="element-with-table-label other-element-with-table-label">';
         atividadesSecund += '<tr style="background-color:#eeeeee;" class="slds-line-height_reset">'+
                             '<th>Código</th><th style="align:left">Descrição</th></tr>';
         for(Atividades_secundarias atvSec : dados.atividades_secundarias) {
             atividadesSecund += '<tr> <td>' +  atvSec.code + '</td><td> ' + atvSec.text + '</td></tr>';
         }
         atividadesSecund += '</table>';
         wrapper.dadosAtividadeSec = atividadesSecund; 
   
         String quadSocietario= '<table class="slds-table slds-table_cell-buffer slds-table_bordered" aria-labelledby="element-with-table-label other-element-with-table-label">';
         quadSocietario += '<tr style="background-color:#eeeeee;" class="slds-line-height_reset">'+
                           '<th class="" scope="col"><div class="slds-truncate" title="Nome">Nome</div></th>'+
                           '<th class="" scope="col"><div class="slds-truncate" title="Qualificação">Qualificação</div></th>'+
                           '<th class="" scope="col"><div class="slds-truncate" title="País de Origem">País</div></th>'+
                           '<th class="" scope="col"><div class="slds-truncate" title="Nome Representante">Representante</div></th>'+
                           '<th class="" scope="col"><div class="slds-truncate" title="Qualificação">Qualificação</div></th>'+
                           '</tr>';
         for (Qsa socio: dados.qsa){
           quadSocietario +='<tr> <td>' + socio.nome + '</td><td> ' + socio.qual + '</td><td> ' + socio.pais_origem + 
                             '</td><td> ' + socio.nome_rep_legal + '</td><td> ' +  socio.qual_rep_legal + '</tr>';
         }
         quadSocietario +='</table>';
         wrapper.dadosQuadSocietario = quadSocietario.replace('null', ''); 
         wrapper.erro = null;
       }else if(response.getStatusCode() == 429){
         wrapper.erro = 'Foram executadas muitas requisições. Espere um minuto para executar novamente.';
       }else if(response.getStatusCode() == 404){
         wrapper.erro = 'Limites de consulta excedido';
       }else if(response.getStatusCode() == 504){
         wrapper.erro = 'TimeOut';
       } else{
         wrapper.erro = response.getStatus();
       }
       
       return wrapper;       
     }    
      
     @AuraEnabled 
     public static void atualizarObjeto( String  recordId2, DadosReceitaWS registro, String dadosQuadSocietario, String atvEconomicaPrincipal, String atvEconomicaSecundaria) {
   
        registro.ultima_atualizacao = Datetime.valueOf(registro.ultima_atualizacao.replace('T',' ')).format();
         Account acc = [ SELECT Id FROM Account WHERE Id = :recordId2 ];
         acc.RazaoSocial__c = registro.nome;
         acc.NomeFantasia__c   = registro.fantasia;
         acc.QuadroSocietario__c = dadosQuadSocietario;
         acc.AtividadeEconomicaSecundaria__c = atvEconomicaSecundaria;
         acc.AtividadeEconomicaPrincipal__c = atvEconomicaPrincipal;
         acc.Bairro__c = registro.bairro;
         acc.CEP__c = registro.cep;
         acc.Cidade__c = registro.municipio;
         acc.Estado__c = registro.uf;
         acc.Numero__c = Decimal.valueOf(registro.numero);
         //acc.Pais__c = 'Brasil';
         acc.Rua__c = registro.logradouro;
         acc.Complemento__c = registro.complemento;
         acc.Telefone__c = registro.telefone;
         acc.CapitalSocial__c = registro.capital_social;
         acc.SituacaoCadastral__c = registro.situacao;
         acc.DataSituacao__c = registro.data_situacao;
         acc.MotivoSituacaoCadastral__c = registro.motivo_situacao;
         acc.SituacaoEspecial__c = registro.situacao_especial;
         acc.DataSituacaEspecial__c = registro.data_situacao_especial;
         acc.DataAbertura__c = registro.abertura;
         acc.TipoEmpresa__c = registro.tipo;
         acc.DataUltimaAtualizacao__c = registro.ultima_atualizacao;
         acc.porte__c = registro.porte;
         acc.NaturezaJuridica__c = registro.natureza_juridica;
         acc.email__c = registro.email;
         acc.EntidadeFederativa__c = registro.efr;
   
         update acc;
     }
   
     @AuraEnabled(cacheable=false)
     public static Sobject consultarDadosObjeto(String idObjeto) {
   
       Account  acc = [  SELECT    CNPJ__c,
                                   RazaoSocial__c,
                                   NomeFantasia__c,
                                   QuadroSocietario__c,
                                   AtividadeEconomicaSecundaria__c,
                                   AtividadeEconomicaPrincipal__c,
                                   Bairro__c,
                                   CEP__c,
                                   Cidade__c,
                                   Estado__c,
                                   Numero__c,
                                   Rua__c,
                                   Complemento__c,
                                   Telefone__c,
                                   CapitalSocial__c,
                                   SituacaoCadastral__c,
                                   DataSituacao__c,
                                   MotivoSituacaoCadastral__c,
                                   SituacaoEspecial__c,
                                   DataSituacaEspecial__c,
                                   DataAbertura__c,
                                   TipoEmpresa__c,
                                   DataUltimaAtualizacao__c,
                                   Porte__c,
                                   NaturezaJuridica__c,
                                   email__c,
                                   EntidadeFederativa__c
                         FROM      Account 
                         WHERE     Id =: idObjeto];
         SObject retorno = (Sobject)acc;
         return retorno;
     }
   
   
     public class Wrapper {
   
       @AuraEnabled public DadosReceitaWS  dadosGeral;
       @AuraEnabled public String dadosAtividadePrim;
       @AuraEnabled public String dadosAtividadeSec;
       @AuraEnabled public String dadosQuadSocietario;
       @AuraEnabled public String erro;
   
     }
   
     public class DadosReceitaWS {
       
       @AuraEnabled  public String ultima_atualizacao {get;set;}
       @AuraEnabled  public String cnpj {get;set;}       
       @AuraEnabled  public String tipo {get;set;}       
       @AuraEnabled  public String porte {get;set;} 
       @AuraEnabled  public String nome {get;set;}   // Razão social
       @AuraEnabled  public String fantasia {get;set;}
       @AuraEnabled  public String abertura {get;set;}   
       @AuraEnabled  public List<Atividade_principal> atividade_principal {get;set;}         
       @AuraEnabled  public List<Atividades_secundarias> atividades_secundarias {get;set;}  
       @AuraEnabled  public String natureza_juridica {get;set;} 
       @AuraEnabled  public String logradouro {get;set;}       
       @AuraEnabled  public String numero {get;set;}           
       @AuraEnabled  public String complemento {get;set;}      
       @AuraEnabled  public String cep {get;set;}              
       @AuraEnabled  public String bairro {get;set;}           
       @AuraEnabled  public String municipio {get;set;}        
       @AuraEnabled  public String uf {get;set;}               
       @AuraEnabled  public String email {get;set;}
       @AuraEnabled  public String telefone {get;set;}         
       @AuraEnabled  public String efr {get;set;}
       @AuraEnabled  public String situacao {get;set;}         
       @AuraEnabled  public String data_situacao {get;set;}    
       @AuraEnabled  public String motivo_situacao {get;set;}   
       @AuraEnabled  public String situacao_especial {get;set;} 
       @AuraEnabled  public String data_situacao_especial {get;set;} 
       @AuraEnabled  public String capital_social {get;set;} 
       @AuraEnabled  public List<Qsa> qsa {get;set;}
   
     }
   
     public class Atividade_principal {
   
       @AuraEnabled  public String code {get;set;} 
       @AuraEnabled  public String text {get;set;} 
   
     }
   
     public class Atividades_secundarias {
   
       @AuraEnabled  public String code {get;set;} 
       @AuraEnabled  public String text {get;set;} 
   
     }
   
     public class Qsa {
       @AuraEnabled  public String nome {get;set;} 
       @AuraEnabled  public String qual {get;set;} 
       @AuraEnabled  public String pais_origem {get;set;} 
       @AuraEnabled  public String nome_rep_legal {get;set;} 
       @AuraEnabled  public String qual_rep_legal {get;set;} 
     }
   
   }