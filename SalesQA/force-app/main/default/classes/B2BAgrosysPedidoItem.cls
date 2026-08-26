public without sharing class B2BAgrosysPedidoItem extends B2BAgrosys {

    public class TpEnvio extends B2BAgrosys.TpEnvio {
        public Integer pro_codigo;
    }

    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public DS ds;
    }

    public B2BAgrosysRetorno calcularImposto(String accountId, String cartId) {

        Account acc = [SELECT Id, CodigoClienteERP__c, Unidade__c, Empresa__c, Lista__c, Ocorrencia__c
                       FROM Account WHERE Id = :accountId];
        
        List<CartItem> itens = [SELECT Id, Name, Quantity, Sku, TotalPrice, ListPrice, TotalAmount, SalesPrice, Product2.PesoUnitario__c, Product2.UnidadeMedida__c, Product2.CodigoProdutoERP__c
                                FROM CartItem
                                WHERE CartId = :cartId];

        Set<String> setUnidadeProdutoKey = new Set<String>();
        for (CartItem item : itens) {
            setUnidadeProdutoKey.add(acc.Unidade__c+'-'+item.Product2.CodigoProdutoERP__c);
        }

        Map<String, Boolean> mapKeyxMultiplicarQuantidade = new Map<String, Boolean>();
        for (UnidadeProduto__c unidade : [SELECT Key__c, MultiplicarQuantidePorPeso__c FROM UnidadeProduto__c WHERE Key__c IN :setUnidadeProdutoKey]) {
            mapKeyxMultiplicarQuantidade.put(unidade.Key__c, unidade.MultiplicarQuantidePorPeso__c);
        }

        dados = new Request();
        dados.request = new Requisicao();
        dados.request.vdadosParam = new List<DadosParam>();
        dados.request.vdadosReceb = new List<dadosReceb>();

        DadosParam dadosParam = new DadosParam();
        dadosParam.tip_operacao = 'SET-PEDIDO';
        dados.request.vdadosParam.add(dadosParam);

        DadosReceb dadosReceb = new DadosReceb();
        dadosReceb.ds = new DS();
        dadosReceb.ds.tp_pedido = new List<TpPedido>();

        TpPedido tpPedido = new TpPedido();
        tpPedido.emp_empresa = Integer.valueOf(acc.Empresa__c);
        tpPedido.uni_unidade = Integer.valueOf(acc.Unidade__c);
        tpPedido.ped_codigo = '';
        tpPedido.ped_palm = '';
        tpPedido.cli_codigo = Integer.valueOf(acc.CodigoClienteERP__c);
        tpPedido.ped_emissao = Date.today();
        tpPedido.ped_entrega = null;
        tpPedido.cpg_condpag = '';
        tpPedido.ped_cashback = 0;
        tpPedido.ped_observ = '';
        tpPedido.ped_obsnf = '';
        // tpPedido.ped_encerrado = false;
        tpPedido.ped_id = '';
        tpPedido.ped_operacao = 1; // Default value for tax calculation
        tpPedido.ind_underline = true;
        tpPedido.tp_vencparc = new List<TpVencparc>();

        tpPedido.tp_pedidoitem = new List<TpPedidoItem>();
        for(CartItem item : itens) {
            TpPedidoItem tpItem = new TpPedidoItem();
            Decimal unitPrice = mapKeyxMultiplicarQuantidade.get(acc.Unidade__c+'-'+item.Product2.CodigoProdutoERP__c) ? item.SalesPrice / item.Product2.PesoUnitario__c : item.SalesPrice;
            tpItem.emp_empresa = Integer.valueOf(acc.Empresa__c);
            tpItem.uni_unidade = Integer.valueOf(acc.Unidade__c);
            tpItem.ped_codigo = '';
            tpItem.pro_codigo = Integer.valueOf(item.Sku);
            tpItem.med_codigo = item.Product2.UnidadeMedida__c;
            tpItem.pit_qtpedi = item.Quantity;
            tpItem.pit_pspedi =  item.Quantity * item.Product2.PesoUnitario__c;
            tpItem.pit_vlpreco = unitPrice;
            tpItem.pit_seq = 0;
            tpItem.pit_observa = '';
            tpItem.pre_tiporeg = (Integer) acc.Lista__c;
            tpitem.pit_tabpre = (Integer) acc.Ocorrencia__c;
            tpItem.pit_status = 0;
            tpItem.pit_ordcomp = '';
            tpItem.pit_seqordcomp = 0;
            tpItem.pit_vlprecoaux = unitPrice;
            tpItem.pit_vlrst = 0;
            tpItem.pit_vfcpst = 0;
            tpItem.pit_vfcpst = 0;

            tpPedido.tp_pedidoitem.add(tpItem);
        }
        dadosReceb.ds.tp_pedido.add(tpPedido);
        dados.request.vdadosReceb.add(dadosReceb);

        String jsonBody = System.JSON.serializePretty(dados);
        jsonBody = jsonBody.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');

        HttpParameters parametro = new HttpParameters();
        parametro.body = jsonBody;
        parametro.method = 'POST';
        parametro.cartId = cartId;
        parametro.accountId = acc.Id;
        parametro.classMethod = 'calcularImposto';
        parametro.className = 'B2BAgrosysPedidoItem';
        parametro.customerCode = acc.CodigoClienteERP__c;
        parametro.integrationName = 'Cálculo de Impostos';
        parametro.cartId = cartId != null ? cartId : null;

        return String.isNotBlank(parametro.body) ? postRequest(parametro) : null;
    }
    
}