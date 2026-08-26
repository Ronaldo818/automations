public with sharing class CheckinCheckoutMapController {

    // Campos custom em Activity (herdados por Event). Lat/Lon são sub-componentes
    // dos campos Geolocation LatitudeLongitudeCheckin__c / LatitudeLongitudeCheckout__c.
    private static final String F_HORA_CHECKIN  = 'HoraCheckin__c';
    private static final String F_LAT_CHECKIN   = 'LatitudeLongitudeCheckin__Latitude__s';
    private static final String F_LON_CHECKIN   = 'LatitudeLongitudeCheckin__Longitude__s';
    private static final String F_HORA_CHECKOUT = 'HoraCheckout__c';
    private static final String F_LAT_CHECKOUT  = 'LatitudeLongitudeCheckout__Latitude__s';
    private static final String F_LON_CHECKOUT  = 'LatitudeLongitudeCheckout__Longitude__s';

    public class CheckinCheckoutLocation {
        @AuraEnabled public Decimal  latCheckin;
        @AuraEnabled public Decimal  lonCheckin;
        @AuraEnabled public Datetime horaCheckin;
        @AuraEnabled public Decimal  latCheckout;
        @AuraEnabled public Decimal  lonCheckout;
        @AuraEnabled public Datetime horaCheckout;
    }

    @AuraEnabled(cacheable=true)
    public static CheckinCheckoutLocation getLocations(Id recordId) {
        String soql = 'SELECT Id, '
                    + F_HORA_CHECKIN  + ', ' + F_LAT_CHECKIN  + ', ' + F_LON_CHECKIN  + ', '
                    + F_HORA_CHECKOUT + ', ' + F_LAT_CHECKOUT + ', ' + F_LON_CHECKOUT
                    + ' FROM Event WHERE Id = :recordId LIMIT 1';
        List<SObject> result = Database.query(soql);

        CheckinCheckoutLocation loc = new CheckinCheckoutLocation();
        if (!result.isEmpty()) {
            SObject e = result[0];
            loc.horaCheckin  = (Datetime) e.get(F_HORA_CHECKIN);
            loc.latCheckin   = (Decimal)  e.get(F_LAT_CHECKIN);
            loc.lonCheckin   = (Decimal)  e.get(F_LON_CHECKIN);
            loc.horaCheckout = (Datetime) e.get(F_HORA_CHECKOUT);
            loc.latCheckout  = (Decimal)  e.get(F_LAT_CHECKOUT);
            loc.lonCheckout  = (Decimal)  e.get(F_LON_CHECKOUT);
        }
        return loc;
    }
}