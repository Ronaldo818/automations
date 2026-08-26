public with sharing class BoletosController {

    @AuraEnabled(cacheable=false)
    public static String buscarCodigoERP(String accountId) {
        if (String.isBlank(accountId)) {
            throw new AuraHandledException('accountId não informado');
        }

        List<Account> contas = [
            SELECT CodigoClienteERP__c
            FROM Account
            WHERE Id = :accountId
            LIMIT 1
        ];

        if (contas.isEmpty() || String.isBlank(contas[0].CodigoClienteERP__c)) {
            throw new AuraHandledException('Código ERP não encontrado para este cliente');
        }

        return contas[0].CodigoClienteERP__c;
    }

    @AuraEnabled(cacheable=false)
    public static String buscarBoletos(String codigoERP) {
        if (String.isBlank(codigoERP)) {
            throw new AuraHandledException('Código ERP não informado');
        }
        return N8NService.buscarBoletos(codigoERP);
    }

    @AuraEnabled(cacheable=false)
    public static void enviarBoletoEmail(String accountId, String pdfBase64,
                                          String nomeArquivo, String valor,
                                          String vencimento, String referencia) {
        if (String.isBlank(accountId) || String.isBlank(pdfBase64)) {
            throw new AuraHandledException('Dados insuficientes para envio');
        }

        try {
            List<Account> contas = [
                SELECT Id, Name, Email__c
                FROM Account
                WHERE Id = :accountId
                LIMIT 1
            ];

            if (contas.isEmpty()) {
                throw new AuraHandledException('Conta não encontrada.');
            }

            Account acc = contas[0];

            if (String.isBlank(acc.Email__c)) {
                throw new AuraHandledException(
                    'O cliente ' + acc.Name + ' não possui email cadastrado.'
                );
            }

            N8NService.enviarEmail(
                acc.Email__c,
                acc.Name,
                nomeArquivo,
                pdfBase64,
                valor,
                vencimento,
                referencia
            );

        } catch (AuraHandledException e) {
            throw e;
        } catch (Exception e) {
            throw new AuraHandledException('Erro ao enviar email: ' + e.getMessage());
        }
    }
}