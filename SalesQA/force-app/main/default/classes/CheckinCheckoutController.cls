public with sharing class CheckinCheckoutController {

    // Campos custom em Activity (herdados por Event). Lat/Lon são sub-componentes
    // dos campos Geolocation LatitudeLongitudeCheckin__c / LatitudeLongitudeCheckout__c.
    private static final String F_HORA_CHECKIN  = 'HoraCheckin__c';
    private static final String F_LAT_CHECKIN   = 'LatitudeLongitudeCheckin__Latitude__s';
    private static final String F_LON_CHECKIN   = 'LatitudeLongitudeCheckin__Longitude__s';
    private static final String F_HORA_CHECKOUT = 'HoraCheckout__c';
    private static final String F_LAT_CHECKOUT  = 'LatitudeLongitudeCheckout__Latitude__s';
    private static final String F_LON_CHECKOUT  = 'LatitudeLongitudeCheckout__Longitude__s';
    private static final String F_RESULTADO_VISITA = 'ResultadoVisita__c';

    @AuraEnabled(cacheable=true)
    public static Map<String, Boolean> getStatus(Id recordId) {
        String soql = 'SELECT Id, ' + F_HORA_CHECKIN + ', ' + F_HORA_CHECKOUT
                    + ' FROM Event WHERE Id = :recordId LIMIT 1';
        List<SObject> result = Database.query(soql);

        Map<String, Boolean> status = new Map<String, Boolean>{
            'hasCheckin'  => false,
            'hasCheckout' => false
        };

        if (!result.isEmpty()) {
            SObject e = result[0];
            status.put('hasCheckin',  e.get(F_HORA_CHECKIN)  != null);
            status.put('hasCheckout', e.get(F_HORA_CHECKOUT) != null);
        }
        return status;
    }

    @AuraEnabled
    public static void doCheckin(Id recordId, Decimal lat, Decimal lon) {
        Event e = new Event(Id = recordId);
        e.put(F_HORA_CHECKIN, Datetime.now());
        e.put(F_LAT_CHECKIN, lat);
        e.put(F_LON_CHECKIN, lon);
        update e;
    }

    @AuraEnabled
    public static void doCheckout(Id recordId, Decimal lat, Decimal lon) {
        // Regra do gestor: o Resultado da Visita (preenchido direto no campo do registro)
        // é obrigatório antes do checkout.
        if (String.isBlank(resultadoVisitaAtual(recordId))) {
            throw new AuraHandledException('Preencha o Resultado da Visita antes de concluir o checkout.');
        }

        Event e = new Event(Id = recordId);
        e.put(F_HORA_CHECKOUT, Datetime.now());
        e.put(F_LAT_CHECKOUT, lat);
        e.put(F_LON_CHECKOUT, lon);
        update e;
    }

    private static String resultadoVisitaAtual(Id recordId) {
        String soql = 'SELECT ' + F_RESULTADO_VISITA + ' FROM Event WHERE Id = :recordId LIMIT 1';
        List<SObject> rows = Database.query(soql);
        if (rows.isEmpty()) {
            return null;
        }
        Object v = rows[0].get(F_RESULTADO_VISITA);
        return (v == null) ? null : String.valueOf(v);
    }
}