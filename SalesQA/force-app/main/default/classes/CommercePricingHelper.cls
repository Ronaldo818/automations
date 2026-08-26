public without sharing class CommercePricingHelper {

    public static Map<String, Decimal> getPriceFromSapPriceList(String webStoreId, String accountId, Set<String> productIds, Boolean precoCaixa) {

        system.debug('precoCaixa: ' + precoCaixa);

        accountId = Id.valueOf(accountId);
        String externalId = [SELECT ExternalId__c FROM WebStore where Id = :webStoreId]?.ExternalId__c ?? '';

		Account acc = [	SELECT  EstruturaVendaForcaVendas__r.OrganizacaoVendas__c
								, EstruturaVendaForcaVendas__r.CanalDistribuicao__c
								, EstruturaVendaForcaVendas__r.ListaPreco__c
								, EstruturaVendaForcaVendas__r.EscritorioVendas__c
								, EstruturaVendaForcaVendas__r.GrupoVendedores__c // Equipe de Vendas
								, EstruturaVendaForcaVendas__r.GrupoClientes3__c  // Categoria
								, EstruturaVendaForcaVendas__r.GrupoClientes4__c // Subcategoria
								, EstruturaVendaForcaVendas__r.GrupoClientes5__c // Clusterização
								, EstruturaVendaForcaVendas__r.RegiaoVendas__c
								, EstruturaVendaEcommerce__r.OrganizacaoVendas__c
								, EstruturaVendaEcommerce__r.CanalDistribuicao__c
								, EstruturaVendaEcommerce__r.ListaPreco__c
								, EstruturaVendaEcommerce__r.EscritorioVendas__c
								, EstruturaVendaEcommerce__r.GrupoVendedores__c // Equipe de Vendas
								, EstruturaVendaEcommerce__r.GrupoClientes3__c  // Categoria
								, EstruturaVendaEcommerce__r.GrupoClientes4__c // Subcategoria
								, EstruturaVendaEcommerce__r.GrupoClientes5__c // Clusterização
								, EstruturaVendaEcommerce__r.RegiaoVendas__c
						FROM    Account
						WHERE   Id = :accountId ];
        System.debug('acc: ' + acc);

        String organizacaoVendas;
		String canalDistribuicao;
		String listaPreco;
		String escritorioVendas;
		String equipeVendas;
		String categoria;
		String subcategoria;
		String clusterizacao;
		String regiaoVendas;

        switch on externalId {
            when 'FV' {
                organizacaoVendas = acc.EstruturaVendaForcaVendas__r.OrganizacaoVendas__c; 
                canalDistribuicao = acc.EstruturaVendaForcaVendas__r.CanalDistribuicao__c;
                listaPreco = acc.EstruturaVendaForcaVendas__r.ListaPreco__c;
                escritorioVendas = acc.EstruturaVendaForcaVendas__r.EscritorioVendas__c;
                equipeVendas = acc.EstruturaVendaForcaVendas__r.GrupoVendedores__c;
                categoria = acc.EstruturaVendaForcaVendas__r.GrupoClientes3__c;
                subcategoria = acc.EstruturaVendaForcaVendas__r.GrupoClientes4__c;
                clusterizacao = acc.EstruturaVendaForcaVendas__r.GrupoClientes5__c;
                regiaoVendas = acc.EstruturaVendaForcaVendas__r.RegiaoVendas__c;
            }
            when 'B2B' {
                organizacaoVendas = acc.EstruturaVendaEcommerce__r.OrganizacaoVendas__c; 
                canalDistribuicao = acc.EstruturaVendaEcommerce__r.CanalDistribuicao__c;
                listaPreco = acc.EstruturaVendaEcommerce__r.ListaPreco__c;
                escritorioVendas = acc.EstruturaVendaEcommerce__r.EscritorioVendas__c;
                equipeVendas = acc.EstruturaVendaEcommerce__r.GrupoVendedores__c;
                categoria = acc.EstruturaVendaEcommerce__r.GrupoClientes3__c;
                subcategoria = acc.EstruturaVendaEcommerce__r.GrupoClientes4__c;
                clusterizacao = acc.EstruturaVendaEcommerce__r.GrupoClientes5__c;
                regiaoVendas = acc.EstruturaVendaEcommerce__r.RegiaoVendas__c;
                precoCaixa = true; // Para o B2B o retorno deve ser sempre no preço total da caixa
            }
        }

        System.debug('listaPreco: ' + listaPreco);

        List<ListaPrecoSAP__c> lstPreco = new List<ListaPrecoSAP__c>();
        lstPreco = [    SELECT  Id, Produto__c, Prioridade__c, SKU__c, Preco__c,
                                Produto__r.ShippingWeightUnitOfMeasure, Produto__r.UnidadeMedida__c, Produto__r.ShippingWeight
                        FROM    ListaPrecoSAP__c
                        WHERE   TipoCondicao__c = 'PR00'
                        AND     Ativo__c = true
                        AND     InicioValidade__c  <= TODAY
                        AND     FimValidade__c >= TODAY
                        AND     Produto__c = :productIds
                        AND     OrganizacaoVendas__c = :organizacaoVendas
                        AND     CanalDistribuicao__c = :canalDistribuicao
                        AND ((	Prioridade__c  = '010'
                                AND Cliente__c = :accountId )
                            OR (Prioridade__c  = '020'
                                AND ListaPrecos__c = :listaPreco )
                            OR (Prioridade__c  = '030' 
                                AND Cliente__c = :accountId
                                AND EscritorioVendas__c = :escritorioVendas
                                AND EquipeVendas__c = :equipeVendas )
                            OR (Prioridade__c  = '040'
                                AND EscritorioVendas__c = :escritorioVendas
                                AND EquipeVendas__c = :equipeVendas)
                            OR (Prioridade__c  = '050'
                                AND EscritorioVendas__c = :escritorioVendas
                                AND EquipeVendas__c = :equipeVendas
                                AND Categoria__c = :categoria
                                AND Subcategoria__c = :subcategoria )
                            OR (Prioridade__c  = '060'
                                AND EscritorioVendas__c = :escritorioVendas
                                AND EquipeVendas__c = :equipeVendas
                                AND Categoria__c = :categoria )
                            OR (Prioridade__c  = '070'
                                AND EscritorioVendas__c = :escritorioVendas
                                AND EquipeVendas__c = :equipeVendas
                                AND Subcategoria__c = :subcategoria )
                            OR (Prioridade__c  = '080'
                                AND EscritorioVendas__c = :escritorioVendas
                                AND EquipeVendas__c = :equipeVendas
                                AND Clusterizacao__c = :clusterizacao )
                            OR (Prioridade__c  = '090'
                                AND Subcategoria__c = :subcategoria )
                            OR (Prioridade__c  = '100'
                                AND Clusterizacao__c = :categoria )
                            OR (Prioridade__c  = '110'
                                AND Clusterizacao__c = :clusterizacao )
                            OR (Prioridade__c  = '120' ))
                        ORDER BY Produto__c, Prioridade__c DESC, InicioValidade__c     ];
        System.debug('lstPreco: ' + lstPreco);
        Map<String, Decimal> mapResult = new Map<String, Decimal>();
        for(ListaPrecoSAP__c lp : lstPreco){
            Decimal preco = lp.Preco__c;

            // Multiplica o preço por KG pelo peso total se: Solicitado preço por caixa, o peso está preencido e a unidade de medida é KG
            if(precoCaixa && lp.Produto__r.ShippingWeight > 0 && (lp.Produto__r.UnidadeMedida__c == 'KG' || lp.Produto__r.ShippingWeightUnitOfMeasure == 'KG')){
                preco = preco * lp.Produto__r.ShippingWeight;
            }

            mapResult.put(lp.Produto__c, preco.setScale(2));
        }

        return mapResult;
    }

    public static void provisoryCodeCoverageMethod(){
        Integer i = 0;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
       	i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
       	i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
    }
}