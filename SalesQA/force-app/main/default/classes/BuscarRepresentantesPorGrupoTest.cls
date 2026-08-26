@IsTest
private class BuscarRepresentantesPorGrupoTest {

    // ─── Helpers ──────────────────────────────────────────────────────────────

    private static User criarUser(String nome, Boolean ativo) {
        Profile p = [SELECT Id FROM Profile WHERE Name = 'Standard User' LIMIT 1];

        User u = new User(
            LastName          = nome,
            Alias             = nome.left(8),
            Email             = nome.replaceAll(' ', '') + '@test.example.com',
            Username          = nome.replaceAll(' ', '') + '@test.example.salesforce',
            EmailEncodingKey  = 'UTF-8',
            LanguageLocaleKey = 'en_US',
            LocaleSidKey      = 'en_US',
            TimeZoneSidKey    = 'America/Sao_Paulo',
            ProfileId         = p.Id,
            IsActive          = ativo
        );
        insert u;
        return u;
    }

    private static void criarEstrutura(String grupo, String nomeRepresentante) {
        EstruturaVenda__c ev = new EstruturaVenda__c(
            GrupoVendedores__c = grupo
        );
        insert ev;

        Account parceiro = new Account(Name = nomeRepresentante);
        insert parceiro;

        EstruturaParceiro__c ep = new EstruturaParceiro__c(
            CodigoFuncaoParceiro__c  = 'ES',
            EstruturaVendaId__c      = ev.Id,
            CodigoParceiroNegocio__c = parceiro.Id
        );
        insert ep;
    }

    // ─── Cenário 1: retorno com resultado ─────────────────────────────────────

    @IsTest
    static void testComResultado() {
        User u = criarUser('João Silva', true);
        criarEstrutura('204', 'João Silva');

        BuscarRepresentantesPorGrupo.Input inp = new BuscarRepresentantesPorGrupo.Input();
        inp.grupoVendedores = '204';

        Test.startTest();
        List<List<User>> result =
            BuscarRepresentantesPorGrupo.buscarRepresentantes(
                new List<BuscarRepresentantesPorGrupo.Input>{ inp }
            );
        Test.stopTest();

        Assert.areEqual(1, result.size(), 'Deve retornar uma lista por input');
        Assert.areEqual(1, result[0].size(), 'Deve retornar um representante');
        Assert.areEqual(u.Id, result[0][0].Id, 'Id deve bater com o User criado');
        Assert.areEqual('João Silva', result[0][0].Name);
    }

    // ─── Cenário 2: sem resultado (grupo inexistente) ─────────────────────────

    @IsTest
    static void testSemResultado() {
        BuscarRepresentantesPorGrupo.Input inp = new BuscarRepresentantesPorGrupo.Input();
        inp.grupoVendedores = '999';

        Test.startTest();
        List<List<User>> result =
            BuscarRepresentantesPorGrupo.buscarRepresentantes(
                new List<BuscarRepresentantesPorGrupo.Input>{ inp }
            );
        Test.stopTest();

        Assert.areEqual(1, result.size(), 'Deve retornar uma lista mesmo sem dados');
        Assert.isTrue(result[0].isEmpty(), 'Lista interna deve estar vazia');
    }

    // ─── Cenário 3: usuário inativo deve ser ignorado ─────────────────────────

    @IsTest
    static void testUsuarioInativo() {
        criarUser('Maria Inativa', false);
        criarEstrutura('205', 'Maria Inativa');

        BuscarRepresentantesPorGrupo.Input inp = new BuscarRepresentantesPorGrupo.Input();
        inp.grupoVendedores = '205';

        Test.startTest();
        List<List<User>> result =
            BuscarRepresentantesPorGrupo.buscarRepresentantes(
                new List<BuscarRepresentantesPorGrupo.Input>{ inp }
            );
        Test.stopTest();

        Assert.areEqual(1, result.size());
        Assert.isTrue(result[0].isEmpty(), 'Usuário inativo não deve aparecer');
    }

    // ─── Cenário 4: bulk com múltiplos grupos ─────────────────────────────────

    @IsTest
    static void testBulk() {
        User u204 = criarUser('Carlos A', true);
        User u205 = criarUser('Pedro B', true);

        criarEstrutura('204', 'Carlos A');
        criarEstrutura('205', 'Pedro B');

        BuscarRepresentantesPorGrupo.Input i1 = new BuscarRepresentantesPorGrupo.Input();
        i1.grupoVendedores = '204';

        BuscarRepresentantesPorGrupo.Input i2 = new BuscarRepresentantesPorGrupo.Input();
        i2.grupoVendedores = '205';

        BuscarRepresentantesPorGrupo.Input i3 = new BuscarRepresentantesPorGrupo.Input();
        i3.grupoVendedores = '999';

        Test.startTest();
        List<List<User>> result =
            BuscarRepresentantesPorGrupo.buscarRepresentantes(
                new List<BuscarRepresentantesPorGrupo.Input>{ i1, i2, i3 }
            );
        Test.stopTest();

        Assert.areEqual(3, result.size(), 'Deve manter posição por input');
        Assert.areEqual(1, result[0].size(), 'Grupo 204: 1 representante');
        Assert.areEqual(u204.Id, result[0][0].Id);
        Assert.areEqual(1, result[1].size(), 'Grupo 205: 1 representante');
        Assert.areEqual(u205.Id, result[1][0].Id);
        Assert.isTrue(result[2].isEmpty(), 'Grupo 999: sem representantes');
    }

    // ─── Cenário 5: input nulo ────────────────────────────────────────────────

    @IsTest
    static void testInputNulo() {
        BuscarRepresentantesPorGrupo.Input inp = new BuscarRepresentantesPorGrupo.Input();
        inp.grupoVendedores = null;

        Test.startTest();
        List<List<User>> result =
            BuscarRepresentantesPorGrupo.buscarRepresentantes(
                new List<BuscarRepresentantesPorGrupo.Input>{ inp }
            );
        Test.stopTest();

        Assert.areEqual(1, result.size());
        Assert.isTrue(result[0].isEmpty(), 'Grupo nulo retorna lista vazia');
    }
}