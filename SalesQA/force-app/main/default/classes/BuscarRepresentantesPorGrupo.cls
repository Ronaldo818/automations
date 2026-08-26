/**
 * Busca representantes (função 'ES') de EstruturaParceiro__c pelo grupoVendedores
 * e retorna os User records correspondentes.
 *
 * Retorno: List<List<User>> — SObject real, necessário para o Flow expor
 * Name (label) e Id (value) no Collection Choice Set nativamente.
 * Inner classes com @InvocableVariable não funcionam no Collection Choice Set.
 */
public with sharing class BuscarRepresentantesPorGrupo {

    // ─── Input ────────────────────────────────────────────────────────────────

    public class Input {
        @InvocableVariable(required=true label='Grupo de Vendedores')
        public String grupoVendedores;
    }

    // ─── Método invocável ─────────────────────────────────────────────────────

    @InvocableMethod(
        label='Buscar Representantes por Grupo de Vendedores'
        description='Retorna lista de Users ativos representantes do grupo informado'
    )
    public static List<List<User>> buscarRepresentantes(List<Input> inputs) {

        List<List<User>> resultados = new List<List<User>>();

        if (inputs == null || inputs.isEmpty()) {
            return resultados;
        }

        // 1. Coleta grupos únicos não nulos (bulk-safe: uma query para todos os inputs)
        Set<String> grupos = new Set<String>();
        for (Input inp : inputs) {
            if (String.isNotBlank(inp?.grupoVendedores)) {
                grupos.add(inp.grupoVendedores);
            }
        }

        // Map: grupo → nomes únicos de parceiros
        Map<String, Set<String>> nomesPorGrupo = new Map<String, Set<String>>();

        if (!grupos.isEmpty()) {
            for (EstruturaParceiro__c ep : [
                SELECT CodigoParceiroNegocio__r.Name,
                       EstruturaVendaId__r.GrupoVendedores__c
                FROM   EstruturaParceiro__c
                WHERE  CodigoFuncaoParceiro__c = 'ES'
                AND    EstruturaVendaId__r.GrupoVendedores__c IN :grupos
            ]) {
                String grupo = ep.EstruturaVendaId__r?.GrupoVendedores__c;
                String nome  = ep.CodigoParceiroNegocio__r?.Name;

                if (String.isBlank(grupo) || String.isBlank(nome)) continue;

                if (!nomesPorGrupo.containsKey(grupo)) {
                    nomesPorGrupo.put(grupo, new Set<String>());
                }
                nomesPorGrupo.get(grupo).add(nome);
            }
        }

        // 2. Uma única query de User para todos os nomes (bulk-safe)
        Set<String> todosNomes = new Set<String>();
        for (Set<String> nomes : nomesPorGrupo.values()) {
            todosNomes.addAll(nomes);
        }

        // RISCO: nomes duplicados entre Users ativos → pega o primeiro por ordem alfabética.
        // Recomenda-se mapear por campo único (ex.: matrícula) em produção.
        Map<String, User> userPorNome = new Map<String, User>();
        if (!todosNomes.isEmpty()) {
            for (User u : [
                SELECT Id, Name
                FROM   User
                WHERE  Name IN :todosNomes
                AND    IsActive = true
                ORDER BY Name ASC
            ]) {
                if (!userPorNome.containsKey(u.Name)) {
                    userPorNome.put(u.Name, u);
                }
            }
        }

        // 3. Monta resultado na mesma ordem dos inputs
        for (Input inp : inputs) {
            List<User> lista = new List<User>();

            if (String.isNotBlank(inp?.grupoVendedores)
                && nomesPorGrupo.containsKey(inp.grupoVendedores)) {

                Set<Id> jaAdicionados = new Set<Id>();

                for (String nome : nomesPorGrupo.get(inp.grupoVendedores)) {
                    User u = userPorNome.get(nome);
                    if (u == null) continue;
                    if (jaAdicionados.contains(u.Id)) continue;

                    lista.add(u);
                    jaAdicionados.add(u.Id);
                }
            }

            resultados.add(lista);
        }

        return resultados;
    }
}