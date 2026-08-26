public without sharing class B2BAgrosysPedido extends B2BAgrosys {
    
    public class DadosReceb extends B2BAgrosys.DadosReceb {
        public DS ds;
    }

    @AuraEnabled
    public static String setPedido(String jsonString) {
        Map<String, String> mapOrderStatus = new Map<String, String> {
            '0' => 'Draft',
            '1' => 'Tax',
            '2' => 'Activated',
            '3' => 'reserved',
            '4' => 'Canceled'
        };

        SendOrderParameter sendOrder = (SendOrderParameter) System.JSON.deserialize(jsonString, B2BAgrosysPedido.SendOrderParameter.class);

        RetornoWrapper retornoWrapper = new RetornoWrapper();
        Account conta;

        Boolean pedidoShouldActivate = false;

        try {
            Order pedido = [
                SELECT Id, Status, GrandTotalAmount, ValorCashback__c, CreatedDate, OrderNumber, NumeroPedidoAgrosys__c, Pricebook2.CodigoTabela__c, PagSeguroPaymentReferenceNumber__c, BandeiraCartao__c, ChavePix__c, AccountId
                FROM Order
                WHERE Id = :sendOrder.orderId
            ];
            pedidoShouldActivate = pedido.Status == 'Draft';
    
            conta = [
                SELECT Id, Unidade__c, Empresa__c, CodigoClienteERP__c, CondicaoPagamento__c, CondicaoPagamento__r.CodigoCondicaoPagamento__c, Ocorrencia__c, Lista__c
                FROM Account
                WHERE Id = :pedido.AccountId
            ];

            retornoWrapper = setRequisicao(pedido, conta, SendOrder);
            if(retornoWrapper.hasError || String.isBlank(retornoWrapper.json)) {
                return System.JSON.serialize(retornoWrapper);
            }

            HttpParameters parametro = new HttpParameters();
            parametro.body = retornoWrapper.json;
            parametro.method = 'POST';
            parametro.classMethod = 'setPedido';
            parametro.className = 'B2BAgrosysPedido';
            parametro.customerCode = conta.CodigoClienteERP__c;
            parametro.orderId = pedido.Id;
            parametro.accountId = conta.Id;
            parametro.integrationName = 'Pedido';
    
            B2BAgrosysRetorno retorno = new B2BAgrosysRetorno();
            retorno = postRequest(parametro);
            retornoWrapper.json = null;

            List<Order> lstOrderUpsert = new List<Order>();

            System.debug('retorno != null => ' + retorno != null);
            System.debug('retorno.hasError => ' + retorno.hasError);
            System.debug('retorno.errorMessage => ' + retorno.errorMessage);
            
            if(retorno != null && retorno.hasError) {
                retornoWrapper.hasError     = true;
                retornoWrapper.errorMessage  = retorno.errorMessage;
                retornoWrapper.logId        = retorno.logId;
                return System.JSON.serialize(retornoWrapper);
            }
            if(retorno != null && retorno.tp_retorno[0]?.msg_codigo != 1) {
                retornoWrapper.hasError     = true;
                retornoWrapper.errorMessage  = retorno.tp_retorno[0]?.msg_descricao;
                retornoWrapper.logId        = retorno.logId;
                return System.JSON.serialize(retornoWrapper);
            }

            if( retorno != null
                && retorno.tp_retorno != null
                && !retorno.tp_retorno.isEmpty()
                && retorno.tp_retorno[0]?.msg_codigo == 1) {    
                        pedido.PedidoEnviadoERP__c  = true;
                        pedido.MetodoPagamento__c   = SendOrder.formaPagamento;
                        pedido.Status               = mapOrderStatus.get(String.valueOf(sendOrder.operacao));

                        if(SendOrder.formaPagamento.equalsIgnoreCase('cartao de credito')) {
                            pedido.NumeroParcelas__c = SendOrder.parcelas;
                        }
                        else if(SendOrder.formaPagamento.equalsIgnoreCase('boleto')) {
                            pedido.NumeroParcelas__c = '1x';
                        }
                        else if(SendOrder.formaPagamento.equalsIgnoreCase('pix')) {
                            pedido.NumeroParcelas__c = '1x';
                            // pedido.URLPix__c = sendOrder.url;
                        }
                        lstOrderUpsert.add(pedido);
            }
            else {
                if(retorno.tp_retorno[0].msg_codigo != 1) {
                    retornoWrapper.hasError = true;
                    List<MessageErrorAgrosys__mdt> lstMessage = [
                        SELECT MessageCode__c, MessageDescription__c
                        FROM MessageErrorAgrosys__mdt
                        WHERE MessageCode__c = :retorno.tp_retorno[0].msg_codigo
                    ];
                    if(!lstMessage.isEmpty()) {
                        retornoWrapper.errorMessage = lstMessage[0].MessageDescription__c;
                    }
                    else {
                        MessageErrorAgrosys__mdt defaultErrorMessage = [SELECT MessageCode__c, MessageDescription__c
                                                                        FROM MessageErrorAgrosys__mdt
                                                                        WHERE MessageCode__c = 999];
                        retornoWrapper.errorMessage = defaultErrorMessage.MessageDescription__c;
                    }

                    retornoWrapper.logId = retorno.logId;
                    return System.JSON.serialize(retornoWrapper);
                }
            }
    
            if(!lstOrderUpsert.isEmpty()) {
                List<IntegrationLog__c> lstLogInsert = new List<IntegrationLog__c>();
                Integer i = 0;
                Database.UpsertResult[] results = Database.upsert(lstOrderUpsert, Order.Id, false);
                for(Database.UpsertResult sr : results){
                    if(!sr.success) {
                        for(Database.Error err : sr.getErrors()) {
                            IntegrationLog__c log = new IntegrationLog__c();
                            log.NomeLog__c = conta != null ? 'ERRO Pedido Agrosys - ' + conta.CodigoClienteERP__c : 'ERRO Pedido Agrosys';
                            log.AccountId__c = conta.Id;
                            log.OrderId__c = sendOrder.orderId;
                            log.ApexClass__c = 'B2BAgrosysPedido';
                            log.ApexMethod__c = 'setPedido';
                            log.OperationStatus__c = 'Error';
                            log.MessageError__c = 'Status: ' + err.getStatusCode() + '\n\nCampos: ' + err.getFields() + '\n\nMensagem' + err.getMessage();                                
                            
                            lstLogInsert.add(log);
                        }
                    }
                    i++;   
                }

                if(!lstLogInsert.isEmpty()) {
                    insert lstLogInsert;
                    retornoWrapper.hasError = true;
                    retornoWrapper.errorMessage = 'Ocorreu um erro ao tentar criar o seu pedido. Por favor, tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema junto com o ID do erro: ' + lstLogInsert[0]?.Id;
                    return System.JSON.serialize(retornoWrapper);
                }
                else {
                    if(Test.isRunningTest() && SendOrder.parcelas.equalsIgnoreCase('error setpedido')) {
            			insert new Account();   
            		}
                    if(pedidoShouldActivate){
                        B2BLWRCheckoutControler.createOrderSummary(lstOrderUpsert[0].Id);
                    }
                    return System.JSON.serialize(retornoWrapper);
                }
            }

            retornoWrapper.hasError = true;
            retornoWrapper.errorMessage = 'Ocorreu um erro ao tentar criar o seu pedido. Por favor, tente novamente. Se o erro persistir informe ao representante de vendas ou administrador do sistema.';
            return System.JSON.serialize(retornoWrapper);
        } catch(Exception ex) {
            retornoWrapper.hasError = true;
            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c = conta != null ? 'ERRO Pedido Agrosys - ' + conta.CodigoClienteERP__c : 'ERRO Pedido Agrosys';
            log.AccountId__c = conta.Id;
            log.OrderId__c = SendOrder.orderId;
			log.ApexClass__c = 'B2BAgrosysPedido';
            log.ApexMethod__c = 'setPedido';
			log.OperationStatus__c = 'Error';
            log.MessageError__c = 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
			insert log;

            System.debug('Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage());

            retornoWrapper.logId = log.Id;
            retornoWrapper.errorMessage = 'Infelizmente seu pedido não pode ser concluído. Informe ao seu representante de vendas ou ao administrador do sistema.';
            return System.JSON.serialize(retornoWrapper);
        }
    }

    public static RetornoWrapper setRequisicao(Order order, Account account, B2BAgrosysPedido.SendOrderParameter orderParameter) {
        Map<String, Integer> mapBrandNameXBrandCode = new Map<String, Integer> {
            'Visa' => 1,
            'MasterCard' => 2
            // 'American Express' => 'American Express',
            // 'Cabal' => 'Cabal',
            // 'Elo' => 'Elo',
            // 'HiperCard' => 'HiperCard',
            // 'Diners Club' => 'Diners Club',
            // 'Hiper' => 'Hiper'
        };
        RetornoWrapper retornoWrapper = new RetornoWrapper();

        try {

            List<OrderItem> itemPedido = [SELECT Id, ValorDesconto__c, ValorST__c, ListPrice, FundoCombatePobrezaST__c, LineNumber, TotalPrice, UnitPrice, Quantity, 
                                          Product2.PesoUnitario__c, UnidadeMedida__c, Product2.StockKeepingUnit, Product2.CodigoProdutoERP__c, TotalAdjustmentAmount
                                          FROM OrderItem WHERE OrderId =: order.Id];

            Set<String> setUnidadeProdutoKey = new Set<String>();
            for (OrderItem item : itemPedido) {
                setUnidadeProdutoKey.add(account.Unidade__c+'-'+item.Product2.CodigoProdutoERP__c);
            }

            Map<String, Boolean> mapKeyxMultiplicarQuantidade = new Map<String, Boolean>();
            for (UnidadeProduto__c unidade : [SELECT Key__c, MultiplicarQuantidePorPeso__c FROM UnidadeProduto__c WHERE Key__c IN :setUnidadeProdutoKey]) {
                mapKeyxMultiplicarQuantidade.put(unidade.Key__c, unidade.MultiplicarQuantidePorPeso__c);
            }
            Request dados = new Request();
            // dados = new Request();
            dados.request = new Requisicao();
            dados.request.vdadosParam = new List<DadosParam>();
            dados.request.vdadosReceb = new List<dadosReceb>();

            DadosParam dadosParam = new DadosParam();
            dadosParam.tip_operacao = 'SET-PEDIDO';
            dados.request.vdadosParam.add(dadosParam);

            DadosReceb dadosReceb = new DadosReceb();
            dadosReceb.ds = new DS();
            dadosReceb.ds.tp_pedido = new List<B2BAgrosys.TpPedido>();

            TpPedido tpPedido = new TpPedido();

            String pedCodigo = String.valueOf(order.NumeroPedidoAgrosys__c).replaceFirst('^0+', '');
            tpPedido.emp_empresa = Integer.valueOf(account.Empresa__c);
            tpPedido.uni_unidade = Integer.valueOf(account.Unidade__c);
            tpPedido.ped_codigo = String.valueOf(pedCodigo);
            tpPedido.ped_palm = '';
            tpPedido.cli_codigo = Integer.valueOf(account.CodigoClienteERP__c);
            tpPedido.ped_emissao = Date.today();
            tpPedido.ped_entrega = Date.today().addDays(7);
            tpPedido.cpg_condpag = account.CondicaoPagamento__r.CodigoCondicaoPagamento__c;
            tpPedido.ped_cashback = order.ValorCashback__c;
            tpPedido.ped_observ = '';
            tpPedido.ped_obsnf = '';
            // tpPedido.ped_encerrado = true;
            tpPedido.ped_operacao = orderParameter.operacao;
            tpPedido.ped_id = orderParameter.orderId;
            tpPedido.ind_underline = true;

            tpPedido.tp_pedidoitem = new List<TpPedidoItem>();

            for(OrderItem item : itemPedido) {
                TpPedidoItem tpItem = new TpPedidoItem();
                Decimal unitPrice = mapKeyxMultiplicarQuantidade.get(account.Unidade__c+'-'+item.Product2.CodigoProdutoERP__c) ? item.UnitPrice / item.Product2.PesoUnitario__c : item.UnitPrice;
                tpItem.emp_empresa = Integer.valueOf(account.Empresa__c);
                tpItem.uni_unidade = Integer.valueOf(account.Unidade__c);
                tpItem.ped_codigo = pedCodigo;
                tpItem.pro_codigo = Integer.valueOf(item.Product2.StockKeepingUnit);
                tpItem.med_codigo = item.UnidadeMedida__c;
                tpItem.pit_qtpedi = item.Quantity;
                tpItem.pit_pspedi = item.Quantity * item.Product2.PesoUnitario__c;
                tpItem.pit_vlpreco = unitPrice;
                tpItem.pit_seq =  item.LineNumber; // sequencia;
                tpItem.pit_observa = '';
                tpItem.pre_tiporeg = (Integer) account.Lista__c;
                tpItem.pit_tabpre = (Integer) account.Ocorrencia__c;
                tpItem.pit_status = 0; // Quais são os possíveis valores?
                tpItem.pit_ordcomp = ''; // O que seria "Ordem de comprado do Cliente"?
                tpItem.pit_seqordcomp = 0; // O que seria "Sequência do item na OC"?
                tpItem.pit_vlprecoaux = unitPrice;
                tpItem.pit_vldesc = -item.TotalAdjustmentAmount;
                tpItem.pit_vlrst = item.ValorST__c;
                tpItem.pit_vfcpst = item.FundoCombatePobrezaST__c;
                tpPedido.tp_pedidoitem.add(tpItem);
            }

            tpPedido.tp_vencparc = new List<TpVencparc>();
            TpVencparc vencParc = new TpVencparc();
            vencParc.emp_empresa = Integer.valueOf(account.Empresa__c);
            vencParc.uni_unidade = Integer.valueOf(account.Unidade__c);
            vencParc.ped_codigo = Integer.valueOf(pedCodigo);
            vencParc.nrdoc = Integer.valueOf(order.OrderNumber);
		    vencParc.valor = order.GrandTotalAmount - order.ValorCashback__c;
            vencParc.seq = 1;
            vencParc.tipo = 'eCO';
		    vencParc.parcent = 'yes';

            if(orderParameter.formaPagamento.equalsIgnoreCase('cartao de credito')) {
            	vencParc.tppag = 'CCR';
		    	vencParc.nrparcela = Integer.valueOf(orderParameter.parcelas.replaceAll('[^0-9]', ''));
            	vencParc.dtvenc = Date.today().addDays(10);
                vencParc.cnpjcartao = System.Label.CNPJPagSeguro;  
                if(order.PagSeguroPaymentReferenceNumber__c != null) {
                    vencParc.nrautoriz = Decimal.valueOf(order.PagSeguroPaymentReferenceNumber__c);
                }
                if(order.BandeiraCartao__c != null) {
                    vencParc.bandeira = mapBrandNameXBrandCode.get(order.BandeiraCartao__c);
                }
            }
            else if(orderParameter.formaPagamento.equalsIgnoreCase('pix')) {
            	vencParc.tppag = 'Pix';
                vencParc.chvpix = order.ChavePix__c;
            	vencParc.dtvenc = Date.today().addDays(10);
		    	vencParc.nrparcela = 1;
            }
            else if(orderParameter.formaPagamento.equalsIgnoreCase('boleto')) {
                vencParc.tppag = 'Boleto';
            	vencParc.dtvenc = Date.today().addDays(10);
		    	vencParc.nrparcela = 1;
            }
            
            if(Test.isRunningTest() && orderParameter.formaPagamento.equalsIgnoreCase('error setrequisicao')) {
            	insert new Account();   
            }
            
            tpPedido.tp_vencparc.add(vencParc);

            dadosReceb.ds.tp_pedido.add(tpPedido);
            dados.request.vdadosReceb.add(dadosReceb);

            String json = System.JSON.serialize(dados);
            retornoWrapper.json = json.replaceAll('_(?=[^"]*":|[^"]*" :)', '-');
            return retornoWrapper;
        }
        catch(Exception ex) {
            RetornoWrapper.hasError = true;
            IntegrationLog__c log = new IntegrationLog__c();
            log.NomeLog__c = 'ERRO Pedido Agrosys - ' + account.CodigoClienteERP__c;
            log.OrderId__c = order.Id;
            log.AccountId__c = account.Id;
			log.ApexClass__c = 'B2BAgrosysPedido';
            log.ApexMethod__c = 'setRequisicao';
			log.OperationStatus__c = 'Error';
            log.MessageError__c = 'Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage();
			insert log;

            System.debug('Type: ' + ex.getTypeName() + '\n\nCause: ' + ex.getCause() + '\n\nLine Number: ' + ex.getLineNumber() + '\n\nStack Trace: ' + ex.getStackTraceString() + '\n\nMessage: ' + ex.getMessage());

            RetornoWrapper.logId = log.Id;
            RetornoWrapper.errorMessage = 'Infelizmente seu pedido não pode ser concluído. Informe ao seu representante de vendas ou ao administrador do sistema.';
            return RetornoWrapper;
        }
    }

    public class SendOrderParameter {
        public String orderId;
        public String parcelas;
        public String cvvCartao;
        public String anoCartao;
        public String mesCartao;
        public String numeroCartao;
        public String nomeCartao;
        public String formaPagamento;
        public String valorCashback;
        public Integer operacao;
    }
}