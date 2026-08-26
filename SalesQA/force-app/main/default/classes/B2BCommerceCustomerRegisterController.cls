/*
 * @description       : 
 * @author            : Allefy Joaquim
 * @group             : 
 * @last modified on  : 09-21-2023
 * @last modified by  : Allefy Joaquim
*/

public without sharing class B2BCommerceCustomerRegisterController {

    @AuraEnabled
    public static String isValidCNPJ(String cnpj) {
    Wrapper wrapperReturn = new Wrapper();
    cnpj = cnpj != null ? cnpj.replaceAll('[^A-Za-z0-9]', '').toUpperCase() : '';
    wrapperReturn.cnpj = cnpj;

    if (cnpj.length() != 14) {
        wrapperReturn.valid = false;
        return System.JSON.serialize(wrapperReturn);
    }

    if (Pattern.matches('^([A-Z0-9])\\1+$', cnpj)) {
        wrapperReturn.valid = false;
        return System.JSON.serialize(wrapperReturn);
    }

    List<Integer> mult1 = new List<Integer>{5,4,3,2,9,8,7,6,5,4,3,2};
    List<Integer> mult2 = new List<Integer>{6,5,4,3,2,9,8,7,6,5,4,3,2};

    Integer sum1 = 0;
    Integer sum2 = 0;

    for (Integer i = 0; i < 12; i++) {
        Integer value = cnpj.substring(i, i + 1).charAt(0) - 48;
        sum1 += value * mult1[i];
        sum2 += value * mult2[i];
    }

    Integer remainder1 = Math.mod(sum1, 11);
    Integer firstDigit = remainder1 < 2 ? 0 : 11 - remainder1;

    sum2 += firstDigit * mult2[12];
    Integer remainder2 = Math.mod(sum2, 11);
    Integer secondDigit = remainder2 < 2 ? 0 : 11 - remainder2;

    wrapperReturn.valid =
        cnpj.substring(12, 13) == String.valueOf(firstDigit) &&
        cnpj.substring(13, 14) == String.valueOf(secondDigit);

    return System.JSON.serialize(wrapperReturn);

    }

    @AuraEnabled
    public static String verifyCustomer(String cnpj){
        cnpj = cnpj != null ? cnpj.replaceAll('[^A-Za-z0-9]', '').toUpperCase() : '';
        Lead lead;
        Account account;
        
        try{
            lead = [SELECT Id, CNPJ__c FROM Lead WHERE CNPJ__c = :cnpj AND IsConverted = false LIMIT 1];
        } catch(Exception ex) {
            lead = null;
        }
        
        try{
            account = [SELECT Id FROM Account WHERE CNPJ__c = :cnpj];
        } catch(Exception ex) {
            account = null;
        }

        Wrapper wrapperReturn = new Wrapper();
        if(lead != null) {
            wrapperReturn.valid = false;
            wrapperReturn.leadId = lead.Id;
            wrapperReturn.mesage = 'O CNPJ ' + cnpj + ' possui um cadastro em análise. Após a avaliação uma notificação será enviada para o e-mail informado durante o cadastrado com mais informações.';
        }
        else if(account != null) {
            wrapperReturn.valid = false;
            wrapperReturn.accountId = account.Id;
            wrapperReturn.mesage = 'O cliente com o CNPJ ' + cnpj + ' já possui um cadastro no sistema.';
        }
        else {
            wrapperReturn.valid = true;
        }
        
        return System.JSON.serialize(wrapperReturn);
    }

    @AuraEnabled
    public static String createLead(String data) {
        try {
            LeadData leadData = (LeadData) System.JSON.deserialize(data, LeadData.class);
    
            Lead newLead = new Lead();
            newLead.FirstName               = leadData.nomeResponsavel;
            newLead.LastName                = leadData.sobrenomeResponsavel;
            newLead.TelefoneResponsavel__c  = leadData.telefoneResponsavel;
            newLead.Company                 = leadData.razaoSocial;
            newLead.CNPJInput__c = leadData.cnpj != null ? leadData.cnpj.replaceAll('[^A-Za-z0-9]', '').toUpperCase() : null;
            newLead.NomeFantasia__c         = leadData.nomeFantasia;
            newLead.Email                   = leadData.email;
            newLead.Phone                   = leadData.telefone;
            newLead.RamoAtividade__c        = leadData.ramoAtividade;
            newLead.Categoria__c            = leadData.categoria;
            newLead.Subcategoria__c         = leadData.subcategoria;
            newLead.Street                  = leadData.rua;
            newLead.BairroEndereco__c       = leadData.bairro;
            newLead.City                    = leadData.cidade;
            newLead.State                   = leadData.estado;
            newLead.NumeroEndereco__c       = leadData.numero;
            newLead.PostalCode              = leadData.cep;
            newLead.ComplementoEndereco__c  = leadData.complemento;
            
            B2BAgrosysPreCadastro preCadastro = new B2BAgrosysPreCadastro();
            String retorno = preCadastro.setCustomer(newLead);
            if(retorno == 'success') {
                insert newLead;
                return newLead.Id;
            } else {
                return 'erro ao enviar formulário!';
            }
        } catch (Exception ex) {
            return 'Error na criação do Lead!' + ex.getMessage();
        }
    }

    public class LeadData {
        public String nomeResponsavel;
        public String sobrenomeResponsavel;
        public String telefoneResponsavel;
        public String razaoSocial;
        public String cnpj;
        public String nomeFantasia;
        public String email;
        public String telefone;
        public String ramoAtividade;
        public String categoria;
        public String subcategoria;
        public String rua;
        public String bairro;
        public String cidade;
        public String estado;
        public String numero;
        public String cep;
        public String complemento;
    }

    public class Option {
        @AuraEnabled
        public String label;
        @AuraEnabled
        public String value;
    }

    public class InitWrapper {
        @AuraEnabled public String errorMessage;
        @AuraEnabled public List<Option> categoria;
        @AuraEnabled public List<Option> subcategoria;

        public InitWrapper() {
            categoria = new List<Option>();
            subcategoria = new List<Option>();
        }
    }

    class Wrapper {
        Boolean valid;
        String cnpj;
        String leadId;
        String accountId;
        String mesage;
    }
}