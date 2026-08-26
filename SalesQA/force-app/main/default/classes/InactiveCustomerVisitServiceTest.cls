@isTest
private class InactiveCustomerVisitServiceTest {

    private static Integer seq = 0;

    private static void injectParams(Boolean ativo) {
        InactiveCustomerVisitHelper.overrideParam = new ParametroVisita__mdt(
            Ativo__c            = ativo,
            DiasSemCompra__c    = 90,
            CadenciaDias__c     = 15,
            DistribuirEmDias__c = 0,
            TipoEvento__c       = 'Visita - Reativação',
            AssuntoEvento__c    = 'Cliente sem Compra'
        );
    }

    private static Account novoRep() {
        Account rep = new Account(Name = 'Rep Teste ' + (seq++));
        insert rep;
        return rep;
    }

    private static User novoUserComunidade(Id repAccId) {
        Contact c = new Contact(LastName = 'Rep Contato', AccountId = repAccId);
        insert c;

        Profile pf = [SELECT Id FROM Profile WHERE Name = 'Commerce Força de Vendas' LIMIT 1];

        Integer n = seq++;
        Long stamp = System.now().getTime();
        User u = new User(
            ContactId         = c.Id,
            ProfileId         = pf.Id,
            LastName          = 'Rep Teste',
            Alias             = 'rept',
            Email             = 'rep.teste@costafood.qa.test',
            Username          = 'rep.teste.' + stamp + '.' + n + '@costafood.qa.test',
            CommunityNickname = 'rept' + stamp + n,
            TimeZoneSidKey    = 'America/Sao_Paulo',
            LocaleSidKey      = 'pt_BR',
            EmailEncodingKey  = 'UTF-8',
            LanguageLocaleKey = 'pt_BR'
        );
        System.runAs(new User(Id = UserInfo.getUserId())) {
            insert u;
        }
        return u;
    }

    private static Account novoCliente(Id repAccId, Boolean evInvalido, Date ultimoFaturamento) {
        Account cli = new Account(
            Name                     = 'Cliente Teste ' + (seq++),
            AgenteVendaExterno__c    = repAccId,
            DataUltimoFaturamento__c = ultimoFaturamento
        );
        insert cli;

        EstruturaVenda__c ev = new EstruturaVenda__c(
            AccountId__c        = cli.Id,
            CadastroInvalido__c = evInvalido
        );
        insert ev;

        cli.EstruturaVendaForcaVendas__c = ev.Id;
        update cli;
        return cli;
    }

    private static List<InactiveCustomerVisitService.Request> reqs(List<Account> accs) {
        List<InactiveCustomerVisitService.Request> rs = new List<InactiveCustomerVisitService.Request>();
        for (Account a : accs) {
            InactiveCustomerVisitService.Request r = new InactiveCustomerVisitService.Request();
            r.accountId = a.Id;
            rs.add(r);
        }
        return rs;
    }

    private static Integer contaVisitas(Id contaId) {
        return [SELECT COUNT() FROM Event WHERE WhatId = :contaId AND Type = 'Visita - Reativação'];
    }


    @isTest
    static void testHappyPathCriaVisita() {
        injectParams(true);
        Account rep = novoRep();
        User u = novoUserComunidade(rep.Id);
        Account cli = novoCliente(rep.Id, false, Date.today().addDays(-100));

        Test.startTest();
        InactiveCustomerVisitService.run(reqs(new List<Account>{ cli }));
        Test.stopTest();

        List<Event> evs = [
            SELECT OwnerId, WhatId, WhoId, Type, Subject, Description,
                   IsAllDayEvent, ActivityDate, StartDateTime, DurationInMinutes
            FROM Event WHERE WhatId = :cli.Id
        ];
        System.assertEquals(1, evs.size(), 'Deveria criar exatamente 1 visita');
        Event e = evs[0];
        System.assertEquals(u.Id, e.OwnerId, 'Dono deve ser o User do representante');
        System.assertEquals(cli.Id, e.WhatId, 'WhatId deve ser a conta');
        System.assertEquals(null, e.WhoId, 'WhoId deve ser nulo');
        System.assertEquals('Visita - Reativação', e.Type, 'Type diferencia de visita a Lead');
        System.assertEquals('Cliente sem Compra', e.Subject, 'Assunto sucinto vem do CMDT (AssuntoEvento__c)');
        System.assertEquals(
            'Cliente sem compra há mais de 90 dias e sem visita nos últimos 15 dias: ' + cli.Name,
            e.Description, 'Descrição completa com os dias parametrizados + nome do cliente');
        System.assertEquals(false, e.IsAllDayEvent, 'Não deve ser compromisso de dia inteiro');
        System.assertEquals(60, e.DurationInMinutes, 'Deve durar 1 hora');
        System.assertNotEquals(null, e.StartDateTime, 'Deve ter horário de início definido');
        System.assertEquals(8, e.StartDateTime.hour(), 'Visita começa às 08:00 (hora local)');
        System.assertEquals(Date.today(), e.ActivityDate, 'ActivityDate (derivada do StartDateTime) é hoje');
    }

    @isTest
    static void testDedupeNaoDuplica() {
        injectParams(true);
        Account rep = novoRep();
        User u = novoUserComunidade(rep.Id);
        Account cli = novoCliente(rep.Id, false, Date.today().addDays(-100));

        insert new Event(
            OwnerId       = u.Id,
            WhatId        = cli.Id,
            Type          = 'Visita - Reativação',
            Subject       = 'Visita pré-existente',
            IsAllDayEvent = true,
            ActivityDate  = Date.today().addDays(-5)   // dentro da cadência de 15 dias
        );

        Test.startTest();
        InactiveCustomerVisitService.run(reqs(new List<Account>{ cli }));
        Test.stopTest();

        System.assertEquals(1, contaVisitas(cli.Id), 'Não deve criar 2ª visita dentro dos 15 dias');
    }

    @isTest
    static void testSemUserAtivoPula() {
        injectParams(true);
        Account rep = novoRep();                        // rep SEM user de comunidade
        Account cli = novoCliente(rep.Id, false, Date.today().addDays(-100));

        Test.startTest();
        InactiveCustomerVisitService.run(reqs(new List<Account>{ cli }));
        Test.stopTest();

        System.assertEquals(0, contaVisitas(cli.Id), 'Sem User ativo no rep não gera visita');
    }

    @isTest
    static void testCadastroInvalidoForaDoEscopo() {
        injectParams(true);
        Account rep = novoRep();
        Account cli = novoCliente(rep.Id, true, Date.today().addDays(-100));   // EV inválida

        Test.startTest();
        InactiveCustomerVisitService.run(reqs(new List<Account>{ cli }));
        Test.stopTest();

        System.assertEquals(0, contaVisitas(cli.Id), 'Cadastro inválido não entra no escopo');
    }

    @isTest
    static void testFaturamentoRecenteForaDoEscopo() {
        injectParams(true);
        Account rep = novoRep();
        Account cli = novoCliente(rep.Id, false, Date.today().addDays(-10));   // < 90 dias

        Test.startTest();
        InactiveCustomerVisitService.run(reqs(new List<Account>{ cli }));
        Test.stopTest();

        System.assertEquals(0, contaVisitas(cli.Id), 'Quem comprou há menos de 90 dias não entra');
    }

    @isTest
    static void testKillSwitchNaoCria() {
        injectParams(false);   // Ativo = false
        Account rep = novoRep();
        Account cli = novoCliente(rep.Id, false, Date.today().addDays(-100));

        Test.startTest();
        InactiveCustomerVisitService.run(reqs(new List<Account>{ cli }));
        Test.stopTest();

        System.assertEquals(0, contaVisitas(cli.Id), 'Com Ativo=false o job não cria nada');
    }

    @isTest
    static void testEntradaVaziaNaoQuebra() {
        injectParams(true);
        Test.startTest();
        InactiveCustomerVisitService.run(null);
        InactiveCustomerVisitService.run(new List<InactiveCustomerVisitService.Request>());
        InactiveCustomerVisitService.run(new List<InactiveCustomerVisitService.Request>{
            new InactiveCustomerVisitService.Request()   // accountId nulo
        });
        Test.stopTest();
        System.assertEquals(0, [SELECT COUNT() FROM Event], 'Entrada vazia não deve criar nada');
    }

    @isTest
    static void testGetParamDefaultExiste() {
        Test.startTest();
        ParametroVisita__mdt p = InactiveCustomerVisitHelper.getParam();
        Test.stopTest();
        System.assertNotEquals(null, p, 'O registro Default do CMDT deve existir no org');
    }
}