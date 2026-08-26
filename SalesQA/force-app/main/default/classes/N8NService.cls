public with sharing class N8NService {

    private static final String WEBHOOK_BOLETOS =
        'https://n8n.avivar.com.br/webhook/45047c80-80a0-4dd9-a5fc-ac3d1f199d90';

    private static final String WEBHOOK_EMAIL =
        'https://n8n.avivar.com.br/webhook/envio-boleto-email';

    private static final String AUTH =
        'Basic QXZpdmFyX3NtMjpBdml2YXJfc20yIUAj';

    public static String buscarBoletos(String codigoERP) {
        if (String.isBlank(codigoERP)) {
            throw new AuraHandledException('Código ERP não informado');
        }

        try {
            HttpRequest req = new HttpRequest();
            req.setEndpoint(WEBHOOK_BOLETOS);
            req.setMethod('POST');
            req.setTimeout(30000);
            req.setHeader('Content-Type', 'application/json');
            req.setHeader('Authorization', AUTH);
            req.setBody(JSON.serialize(
                new Map<String, String>{ 'documento' => codigoERP }
            ));

            HttpResponse res = new Http().send(req);

            if (res.getStatusCode() != 200) {
                throw new AuraHandledException(
                    'Webhook retornou status ' + res.getStatusCode()
                );
            }

            return res.getBody();

        } catch (AuraHandledException e) {
            throw e;
        } catch (Exception e) {
            throw new AuraHandledException('Erro no callout: ' + e.getMessage());
        }
    }

    public static String enviarEmail(String email, String nome,
                                      String nomeArquivo, String pdfBase64,
                                      String valor, String vencimento,
                                      String referencia) {
        try {
            Map<String, String> payload = new Map<String, String>{
                'email'       => email,
                'nome'        => nome,
                'nomeArquivo' => nomeArquivo,
                'pdf'         => pdfBase64,
                'valor'       => valor,
                'vencimento'  => vencimento,
                'referencia'  => referencia
            };

            HttpRequest req = new HttpRequest();
            req.setEndpoint(WEBHOOK_EMAIL);
            req.setMethod('POST');
            req.setTimeout(30000);
            req.setHeader('Content-Type', 'application/json');
            req.setHeader('Authorization', AUTH);
            req.setBody(JSON.serialize(payload));

            HttpResponse res = new Http().send(req);

            System.debug('Email webhook status: ' + res.getStatusCode());
            System.debug('Email webhook body: ' + res.getBody());

            if (res.getStatusCode() != 200) {
                throw new AuraHandledException(
                    'Erro no envio: status ' + res.getStatusCode()
                );
            }

            return res.getBody();

        } catch (AuraHandledException e) {
            throw e;
        } catch (Exception e) {
            throw new AuraHandledException('Erro no callout de email: ' + e.getMessage());
        }
    }
}