public without sharing class InactiveCustomerVisitService {

    public class Request {
        @InvocableVariable(label='ID da Conta' required=true)
        public Id accountId;
    }

    @InvocableMethod(
        label='Gerar Visita de Reativação'
        description='Gera Events de visita para contas inativas. Recebe um lote de contas do Scheduled Flow.')
    public static void run(List<Request> requests) {
        if (requests == null || requests.isEmpty()) {
            return;
        }

        Set<Id> accountIds = new Set<Id>();
        for (Request r : requests) {
            if (r.accountId != null) {
                accountIds.add(r.accountId);
            }
        }
        if (accountIds.isEmpty()) {
            return;
        }

        ParametroVisita__mdt p = InactiveCustomerVisitHelper.getParam();
        if (p.Ativo__c != true) {
            System.debug(LoggingLevel.INFO, 'Visita reativação: parâmetro Ativo=false — lote ignorado.');
            return;
        }

        Integer diasSemCompra = (Integer) p.DiasSemCompra__c;
        Integer cadencia      = (Integer) p.CadenciaDias__c;
        Integer distribuir    = (p.DistribuirEmDias__c == null) ? 0 : (Integer) p.DistribuirEmDias__c;

        Date hoje             = Date.today();
        Date corteFaturamento = hoje.addDays(-diasSemCompra); 
        Date corteDedupe      = hoje.addDays(-cadencia);

        
        Map<Id, Account> elegiveis = new Map<Id, Account>();
        Set<Id> repAccountIds = new Set<Id>();
        for (Account a : [
            SELECT Id, Name, AgenteVendaExterno__c, DataUltimoFaturamento__c
            FROM Account
            WHERE Id IN :accountIds
              AND AgenteVendaExterno__c != null
              AND DataUltimoFaturamento__c != null
              AND DataUltimoFaturamento__c <= :corteFaturamento
              AND EstruturaVendaForcaVendas__c != null
              AND EstruturaVendaForcaVendas__r.CadastroInvalido__c = false
        ]) {
            elegiveis.put(a.Id, a);
            repAccountIds.add(a.AgenteVendaExterno__c);
        }
        if (elegiveis.isEmpty()) {
            return;
        }

        InactiveCustomerVisitHelper.OwnerResolution owners =
            InactiveCustomerVisitHelper.resolveOwners(repAccountIds);
        Set<Id> jaTemVisita = InactiveCustomerVisitHelper.accountsWithRecentVisit(
            elegiveis.keySet(), p.TipoEvento__c, corteDedupe);

        
        List<Event> aCriar = new List<Event>();
        Integer puladasDuplicada = 0;
        Integer puladasSemDono   = 0;
        for (Account a : elegiveis.values()) {
            if (jaTemVisita.contains(a.Id)) {
                puladasDuplicada++;
                continue;
            }
            Id ownerId = owners.accountToUser.get(a.AgenteVendaExterno__c);
            if (ownerId == null) {
                puladasSemDono++;
                continue; // rep sem Person Contact / sem User ativo
            }
            Date dataVisita = (distribuir > 0)
                ? hoje.addDays(Math.mod(Math.abs(String.valueOf(a.Id).hashCode()), distribuir + 1))
                : hoje;
            aCriar.add(InactiveCustomerVisitHelper.buildVisitEvent(a, ownerId, dataVisita, p));
        }

        
        Integer criadas = 0;
        Integer erros   = 0;
        if (!aCriar.isEmpty()) {
            Database.SaveResult[] results = Database.insert(aCriar, false);
            for (Integer i = 0; i < results.size(); i++) {
                if (results[i].isSuccess()) {
                    criadas++;
                } else {
                    erros++;
                    System.debug(LoggingLevel.ERROR,
                        'Visita reativação: falha ao criar Event p/ conta ' + aCriar[i].WhatId
                        + ' -> ' + results[i].getErrors());
                }
            }
        }

        System.debug(LoggingLevel.INFO, String.format(
            'Visita reativação (lote de {0}): {1} criadas, {2} duplicadas, {3} sem dono, {4} erros.',
            new List<Object>{ accountIds.size(), criadas, puladasDuplicada, puladasSemDono, erros }));
    }
}