public without sharing class InactiveCustomerVisitHelper {

    public class VisitaConfigException extends Exception {}

    @TestVisible
    private static ParametroVisita__mdt overrideParam;

    
    public static ParametroVisita__mdt getParam() {
        if (overrideParam != null) {
            return overrideParam;
        }
        ParametroVisita__mdt p = ParametroVisita__mdt.getInstance('Default');
        if (p == null) {
            throw new VisitaConfigException(
                'Registro Default de ParametroVisita__mdt não encontrado. Configure os parâmetros antes de agendar o job.');
        }
        return p;
    }

    
    public class OwnerResolution {
        public Map<Id, Id> accountToUser     = new Map<Id, Id>();
        public Set<Id>     noUserAccounts    = new Set<Id>();
        public Set<Id>     multiUserAccounts = new Set<Id>();
    }

    
    public static OwnerResolution resolveOwners(Set<Id> repAccountIds) {
        OwnerResolution res = new OwnerResolution();
        if (repAccountIds == null || repAccountIds.isEmpty()) {
            return res;
        }

        // rep Account -> Contato(s)
        Map<Id, Id> repByContact = new Map<Id, Id>();
        for (Contact c : [
            SELECT Id, AccountId
            FROM Contact
            WHERE AccountId IN :repAccountIds
        ]) {
            repByContact.put(c.Id, c.AccountId);
        }
        if (repByContact.isEmpty()) {
            res.noUserAccounts.addAll(repAccountIds);
            return res;
        }

        // Contato(s) -> User(s) ativo(s), já agrupados pelo rep Account
        Map<Id, List<User>> usersByRep = new Map<Id, List<User>>();
        for (User u : [
            SELECT Id, ContactId, LastLoginDate
            FROM User
            WHERE ContactId IN :repByContact.keySet()
              AND IsActive = true
        ]) {
            Id repId = repByContact.get(u.ContactId);
            List<User> bucket = usersByRep.get(repId);
            if (bucket == null) {
                bucket = new List<User>();
                usersByRep.put(repId, bucket);
            }
            bucket.add(u);
        }

        // escolhe o dono de cada rep Account
        for (Id repId : repAccountIds) {
            List<User> users = usersByRep.get(repId);
            if (users == null || users.isEmpty()) {
                res.noUserAccounts.add(repId);
            } else if (users.size() == 1) {
                res.accountToUser.put(repId, users[0].Id);
            } else {
                res.multiUserAccounts.add(repId);
                res.accountToUser.put(repId, pickDeterministic(users));
            }
        }
        return res;
    }

    private static Id pickDeterministic(List<User> users) {
        User best = null;
        for (User u : users) {
            if (best == null || isBetter(u, best)) {
                best = u;
            }
        }
        return best.Id;
    }

    private static Boolean isBetter(User candidate, User current) {
        if (candidate.LastLoginDate != current.LastLoginDate) {
            if (candidate.LastLoginDate == null) return false;
            if (current.LastLoginDate == null)   return true;
            return candidate.LastLoginDate > current.LastLoginDate;
        }
        return String.valueOf(candidate.Id).compareTo(String.valueOf(current.Id)) < 0;
    }

    public static Set<Id> accountsWithRecentVisit(Set<Id> accountIds, String tipo, Date cutoff) {
        Set<Id> jaTem = new Set<Id>();
        if (accountIds == null || accountIds.isEmpty()) {
            return jaTem;
        }
        for (Event e : [
            SELECT WhatId
            FROM Event
            WHERE Type = :tipo
              AND WhatId IN :accountIds
              AND ActivityDate >= :cutoff
        ]) {
            if (e.WhatId != null) {
                jaTem.add(e.WhatId);
            }
        }
        return jaTem;
    }

    // Hora em que a visita começa no dia (08:00); dura 1h => 08:00–09:00.
    private static final Integer HORA_INICIO_VISITA = 8;

    public static Event buildVisitEvent(Account acct, Id ownerId, Date dataVisita, ParametroVisita__mdt p) {
        Event e = new Event();
        e.OwnerId           = ownerId;            // dono = User de comunidade do representante
        e.WhatId            = acct.Id;            // Related To = Account (a loja/cliente)
        e.WhoId             = null;               // visita de reativação não aponta para Lead/Contato
        e.Type              = p.TipoEvento__c;    // diferencia de visita a Lead
        e.Subject           = p.AssuntoEvento__c; // título sucinto p/ a agenda (ex.: "Cliente sem Compra")
        // Descrição completa, montada a partir dos parâmetros (DiasSemCompra/CadenciaDias) + nome do cliente.
        e.Description       = String.format(
            'Cliente sem compra há mais de {0} dias e sem visita nos últimos {1} dias: {2}',
            new List<Object>{
                (Integer) p.DiasSemCompra__c,
                (Integer) p.CadenciaDias__c,
                acct.Name
            });
        e.IsAllDayEvent     = false;              // não é mais compromisso de dia inteiro
        e.StartDateTime     = Datetime.newInstance(
            dataVisita, Time.newInstance(HORA_INICIO_VISITA, 0, 0, 0));
        e.DurationInMinutes = 60;                 // 1 hora (08:00–09:00)
        return e;
    }
}