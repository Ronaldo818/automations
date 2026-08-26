@isTest
private class CheckinCheckoutControllerTest {

    private static Id novoEvento() {
        Event e = new Event(
            Subject           = 'Visita Teste',
            StartDateTime     = Datetime.now(),
            DurationInMinutes = 60
        );
        insert e;
        return e.Id;
    }

    @isTest
    static void testCheckinRegistraHora() {
        Id evId = novoEvento();
        Test.startTest();
        CheckinCheckoutController.doCheckin(evId, 1.23, 4.56);
        Test.stopTest();

        Event e = [SELECT HoraCheckin__c FROM Event WHERE Id = :evId];
        System.assertNotEquals(null, e.HoraCheckin__c, 'Checkin deve gravar a hora');
    }

    @isTest
    static void testCheckoutBloqueiaSemResultado() {
        Id evId = novoEvento();
        CheckinCheckoutController.doCheckin(evId, 1.0, 2.0);   // campo Resultado da Visita vazio

        Boolean bloqueado = false;
        Test.startTest();
        try {
            CheckinCheckoutController.doCheckout(evId, 1.0, 2.0);
        } catch (AuraHandledException ex) {
            bloqueado = true;
        }
        Test.stopTest();

        System.assertEquals(true, bloqueado, 'Checkout sem Resultado da Visita deve ser bloqueado');
        Event e = [SELECT HoraCheckout__c FROM Event WHERE Id = :evId];
        System.assertEquals(null, e.HoraCheckout__c, 'Checkout não pode ter sido gravado');
    }

    @isTest
    static void testCheckoutOkComResultadoPreenchido() {
        Id evId = novoEvento();
        CheckinCheckoutController.doCheckin(evId, 1.0, 2.0);
        // Resultado preenchido direto no campo do registro (como o rep faz no layout)
        update new Event(Id = evId, ResultadoVisita__c = 'Convertida');

        Test.startTest();
        CheckinCheckoutController.doCheckout(evId, 3.0, 4.0);
        Test.stopTest();

        Event e = [SELECT HoraCheckout__c, ResultadoVisita__c FROM Event WHERE Id = :evId];
        System.assertNotEquals(null, e.HoraCheckout__c, 'Checkout deve concluir com o resultado preenchido');
        System.assertEquals('Convertida', e.ResultadoVisita__c, 'Resultado não deve ser alterado pelo checkout');
    }

    @isTest
    static void testGetStatus() {
        Id evId = novoEvento();
        Test.startTest();
        Map<String, Boolean> antes = CheckinCheckoutController.getStatus(evId);
        CheckinCheckoutController.doCheckin(evId, 1.0, 2.0);
        update new Event(Id = evId, ResultadoVisita__c = 'Não convertida');
        CheckinCheckoutController.doCheckout(evId, 3.0, 4.0);
        Map<String, Boolean> depois = CheckinCheckoutController.getStatus(evId);
        Test.stopTest();

        System.assertEquals(false, antes.get('hasCheckin'),  'Antes: sem checkin');
        System.assertEquals(false, antes.get('hasCheckout'), 'Antes: sem checkout');
        System.assertEquals(true,  depois.get('hasCheckin'),  'Depois: com checkin');
        System.assertEquals(true,  depois.get('hasCheckout'), 'Depois: com checkout');
    }
}