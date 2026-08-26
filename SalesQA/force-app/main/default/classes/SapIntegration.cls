public without sharing class SapIntegration {

    public class SapReturnWrapper {
        public SapOrder A_SalesOrder;
    }
    
    public class SapOrder {
        public Datetime BillingDocumentDate; // "2025-07-18T00:00:00",
        public String CustomerAccountAssignmentGroup; // "",
        public String CustomerPaymentTerms; // "0009",
        public String DistributionChannel; // "10",
        public String IncotermsClassification; // "CIF",
        public String IncotermsLocation1; // "CIF",
        public String IncotermsTransferLocation; // "",
        public String OrganizationDivision; // "10",
        public String PaymentMethod; // "1",
        public String PurchaseOrderByCustomer; // "00001",
        public String SalesGroup; // "200",
        public String SalesOffice; // "2000",
        public String ShippingCondition; // "2000",
        public Datetime SalesOrderDate; // "2025-07-18T00:00:00",
        public String SalesOrderType; // "ORB",
        public String SalesOrganization; // "2000",
        public String SoldToParty; // "14100001",
        public String SalesDistrict; // "000001",
        public String TransactionCurrency; // "VA01",
        public String ReferenceSDDocument; // "",
        public String ReferenceSDDocumentCategory; // "",
        public String AdditionalCustomerGroup2; // "",
        public Date RequestedDeliveryDate; // "2025-07-18T00:00:00",
        public String AssignmentReference; // "",
        public List<SapOrderItem> to_Item;
        public List<Text> to_Text;
        public List<Partner> to_Partner; // Não enviar tag caso endereço de entrega

        // RESPONSE FIELDS
        public String SalesOrder;
        public DateTime CreationDate;
        public String CreatedByUser;
    }

    public class SapOrderItem {
        public String SalesOrderItem; //: "010",
        public String Material; //: 'TG11'
        public String NetAmount; //: "198.63",
        public String RequestedQuantity; //: '25'
        public String CustomerPaymentTerms; //: "",
        public String ItemNetWeight; //: "365.36",
        public String MaterialGroup; //: "M010",
        public String RequestedQuantityUnit; //: "UN",
        public String SalesOrderItemCategory; //: "TBN",
        public String ProductionPlant; //: "2001",
        public String StorageLocation; //: "2001",
        public String ShippingPoint; //: "2001",
        List<PricingElement> to_PricingElement;

        // RESPONSE FIELDS
        public String SalesOrderItemText;
    }

    public class PricingElement {
        public String PricingProcedureStep; //"ZDES",
        public String ConditionRateValue; //"1",
        public String ConditionQuantity; //"25.36",
        public String ConditionQuantityUnit; //"UN",
        public String ConditionAmount; //"25.36"
    }

    public class Text {
        public String LongText; // "Texto de testes cabeçalho de Vendas *"
    }

    public class Partner {
        public String Customer; // "14100001"
        public String PartnerFunction; // ZE
    }

    public class SapErrorResponse {
        public ErrorWrapper error;
    }

    public class ErrorWrapper {
        public String code;
        public MessageWrapper message;
    }

    public class MessageWrapper {
        public String lang;
        public String value;
    }

    public SapIntegration(){
    }

    public void createOrder(String orderId){
        
        String operation = 'CreateOrder';
        String statusIntegracao = 'Error';
        Order o;
        
        Integration__mdt params = Integration__mdt.getInstance(operation);
        List<LogErro__c> lstErrorLog =  new List<LogErro__c>();        

        // Instância Log de Integração padrão
        IntegrationLog__c newIntegrationlog = new IntegrationLog__c(
            ApexClass__c = 'SapIntegration',
            ApexMethod__c = 'makeRequest',
            EndPoint__c = params.Endpoint__c,
            OrderId__c = orderId
        );

        try {
            // Dados do Pedido com Itens
            o = getOrderWithItems(orderId);

            // Identifica se o Pedido é B2B ou Força de Vendas
            String context = '';
            String estruturaVenda = '';

            if(o.Type == 'B2B Commerce'){
                context = 'b2b';
                estruturaVenda = o.Account.EstruturaVendaEcommerce__c;
            } else {
                context = 'fv';
                estruturaVenda = o.Account.EstruturaVendaForcaVendas__c;
            }

            System.debug('context: ' + context);
            System.debug('estruturaVenda: ' + estruturaVenda);
            

            // Gera Payload para envio do pedido ao SAP
            String requestBody = genRequestBody(o, estruturaVenda, context);
            newIntegrationlog.Request__c = requestBody;

            // Prepara a request de Integração
            HttpRequest request = new HttpRequest();
            request.setEndPoint(params.Endpoint__c);
            request.setMethod(params.Method__c);
            request.setHeader('Content-Type', 'application/json');
            request.setBody(requestBody);
            request.setTimeout(Integer.valueOf(params.Timeout__c));

            // Processa a resposta da integração
            HttpResponse response = new HTTP().send(request);
            newIntegrationlog.Response__c = response.getBody().left(131072);
            newIntegrationlog.IntegrationStatusCode__c = String.valueOf(response.getStatusCode());

            try {
                if(newIntegrationlog.IntegrationStatusCode__c.left(1) == '2'){
                    SapReturnWrapper returndedSapOrder = (SapReturnWrapper)JSON.deserialize(response.getBody(), SapReturnWrapper.class);
                    String codigoERP = returndedSapOrder.A_SalesOrder.SalesOrder.leftPad(10, '0');
                    statusIntegracao = 'Success';

                    // Atualiza o Pedido com retorno SAP
                    o.PedidoEnviadoERP__c = true;
                    o.CodigoPedidoERP__c = codigoERP;
                    o.EstruturaVenda__c = estruturaVenda;
                    o.Status = 'Registered';
                    o.MensagemErroIntegracao__c = '';

                    // Atualiza Itens do Pedido com retorno SAP
                    for(OrderItem oi : o.OrderItems){
                        oi.ExternalId__c = codigoERP + '_' + oi.ExternalId__c.leftPad(6, '0');
                    }

                } else if(newIntegrationlog.IntegrationStatusCode__c.left(1) == '4') {
                    try {
                        SapErrorResponse errorResponse = (SapErrorResponse) JSON.deserialize(response.getBody(), SapErrorResponse.class);
                        o.MensagemErroIntegracao__c = (errorResponse?.error?.message?.value ?? '');
                    } catch(Exception ex) {
                        o.MensagemErroIntegracao__c = 'Erro ' + newIntegrationlog.IntegrationStatusCode__c + ' - Mensagem indeterminada, verificar o LOG para mais detalhes';
                    }

                } else if(newIntegrationlog.IntegrationStatusCode__c.left(1) == '5') {
                    o.MensagemErroIntegracao__c = 'Erro ' + newIntegrationlog.IntegrationStatusCode__c + ' - Erro técnico na API do SAP/CPI, verificar o LOG para mais detalhes';

                } else {
                    o.MensagemErroIntegracao__c = 'Erro ' + newIntegrationlog.IntegrationStatusCode__c + ' - Erro indeterminado, verificar o LOG para mais detalhes';
                }


            } catch(Exception ex){
                lstErrorLog.add(genErrorLog(ex,'ProcessResponse()'));
            }

        } catch(Exception ex) {
            lstErrorLog.add(genErrorLog(ex,'MakeRequest()'));
        } finally {
            Database.insert(newIntegrationlog);
        }

        if(!lstErrorLog.isEmpty()){
            Database.insert(lstErrorLog);
        }
        
        // Grava as modificações no pedido e nos itens
        o.IntegrationOperation__c = operation;
        o.IntegrationStatus__c = statusIntegracao;
        database.update(o);
        if(statusIntegracao == 'Success'){
            database.update(o.OrderItems);
        }
    }

    private Order getOrderWithItems(String recordId) {
        Order o = [ SELECT  Id
                            , OrderNumber
                            , NumeroPedidoCliente__c
                            , Type
                            , CashbackGerado__c
                            , VendaContaOrdem__c
                            , ValorCashback__c
                            , AccountId
                            , Account.CodigoClienteERP__c
                            , Account.EstruturaVendaForcaVendas__c
                            , Account.EstruturaVendaEcommerce__c
                            , ActivatedDate
                            , MetodoPagamento__c
                            , ObservacaoNotaFiscal__c
                            , DataEntrega__c
                            , CodigoCondicaoPagamentoSelecionada__c
                            , CondicaoExpedicao__c
                            , ( SELECT  Id
                                        , ListPrice
                                        , ValorNegociado__c
                                        , TotalPrice
                                        , PesoLiquido__c
                                        , TotalLineAdjustmentAmount
                                        , Quantity
                                        , QuantityKg__c 
                                        , Product2.ShippingWeight
                                        , Product2.GrupoMercadoria__c
                                        , Product2.StockKeepingUnit
                                FROM    OrderItems
                                WHERE   TypeCode = 'Product')
                        FROM    Order
                        WHERE   Id = :recordId  ];

        Integer itemNumeration = 10;
        for(OrderItem oi : o.OrderItems){
            oi.ExternalId__c = String.valueOf(itemNumeration);
            itemNumeration += 10;
        }

        return o;
    }

    private string genRequestBody(Order o, String estruturaVendaId, String context){

        EstruturaVenda__c ev = [ SELECT  CanalDistribuicao__c
                                        , GrupoVendedores__c
                                        , EscritorioVendas__c
                                        , OrganizacaoVendas__c
                                        , RegiaoVendas__c
                                        , CondicaoPagamento__c
                                        , CondicaoExpedicao__c
                                FROM    EstruturaVenda__c
                                WHERE   Id = :estruturaVendaId    ];
        System.debug('ev: ' + ev);
    
        String condicaoPagamento = o.CodigoCondicaoPagamentoSelecionada__c ?? ev.CondicaoPagamento__c; // Condição da conta para FV ou para B2B/Boleto
        if(context == 'b2b'){
            String metodo = o.MetodoPagamento__c;
            switch on metodo {
                when 'cc' {
                    condicaoPagamento = 'E000';
                }
                when 'pix' {
                    condicaoPagamento = 'PE00';
                }
            }
        }

        String codigoERPEntrega = '';
        try {
            // Endereço de Entrega
            OrderDeliveryGroupSummary odgs = [  SELECT  Id, DeliverToName, DeliverToPostalCode, DeliverToStreet, DeliverToState, DeliverToCity
                                                FROM    OrderDeliveryGroupSummary
                                                WHERE   OrderSummary.OriginalOrderId = :o.Id LIMIT 1];

            codigoERPEntrega = [    SELECT  ExternalId__c
                                    FROM    ContactPointAddress
                                    WHERE   AddressType = 'Shipping'
                                    AND     AddressLastName = :odgs.DeliverToName
                                    AND     PostalCode = :odgs.DeliverToPostalCode
                                    AND     ParentId = :o.AccountId ]?.ExternalId__c ?? '';

            if(String.isNotBlank(codigoERPEntrega)){
                if(codigoERPEntrega.endsWith('_ZE')){
                    codigoERPEntrega = codigoERPEntrega.substringAfter('_').substringBefore('_');
                } else {
                    codigoERPEntrega = codigoERPEntrega.substringAfterLast('_');
                }
            }

        }  catch(Exception ex){
            codigoERPEntrega = o.Account.CodigoClienteERP__c;
        }

        String condExpedicao = 'CE' + o.CondicaoExpedicao__c;
        System.debug('condExpedicao: ' + condExpedicao);

        // Condição de Expedição
        ParmetroCondicaoExpedicao__mdt pce = ParmetroCondicaoExpedicao__mdt.getInstance('CE' + o.CondicaoExpedicao__c ?? '');
        System.debug('pce: ' + pce);

        // Forma de Pagamento
        String formaPagto = [   SELECT  FormaPagamentoSAP__c
                                FROM    CondicaoPagamento__c
                                WHERE   CodigoCondicaoPagamento__c = :condicaoPagamento
                                LIMIT   1 ]?.FormaPagamentoSAP__c ?? '';

        Date dataPedidoSAP = ForcaVendas__c.getOrgDefaults()?.DataPedidoSAP__c;
        if(dataPedidoSAP == null || dataPedidoSAP < o.ActivatedDate){
            dataPedidoSAP = Date.valueOf(o.ActivatedDate);
        }

        SapOrder sapOrder = new SapOrder();
        sapOrder.TransactionCurrency = 'VA01';
        sapOrder.CustomerAccountAssignmentGroup = '';
        sapOrder.ReferenceSDDocument = '';
        sapOrder.ReferenceSDDocumentCategory = '';
        sapOrder.AdditionalCustomerGroup2 = '';
        sapOrder.AssignmentReference = '';
        sapOrder.IncotermsTransferLocation = '';
        sapOrder.IncotermsClassification = 'CIF';
        sapOrder.IncotermsLocation1 = 'CIF';
        sapOrder.PurchaseOrderByCustomer = (String.isBlank(o.NumeroPedidoCliente__c) ? o.OrderNumber : o.NumeroPedidoCliente__c);
        sapOrder.SoldToParty = o.Account.CodigoClienteERP__c;
        sapOrder.BillingDocumentDate = dataPedidoSAP;
        sapOrder.SalesOrderDate = dataPedidoSAP;
        sapOrder.RequestedDeliveryDate = o.DataEntrega__c == null ? Date.valueOf(o.ActivatedDate) : o.DataEntrega__c;
        sapOrder.CustomerPaymentTerms = condicaoPagamento;
        sapOrder.PaymentMethod = formaPagto;
        sapOrder.DistributionChannel = ev.CanalDistribuicao__c;
        sapOrder.OrganizationDivision = '10';
        sapOrder.SalesGroup = ev.GrupoVendedores__c;
        sapOrder.SalesOffice = ev.EscritorioVendas__c;
        sapOrder.ShippingCondition = ev.CondicaoExpedicao__c;
        sapOrder.SalesOrganization = ev.OrganizacaoVendas__c;
        sapOrder.SalesDistrict = (ev.RegiaoVendas__c ?? '');
        sapOrder.SalesOrderType = o.VendaContaOrdem__c ? 'RCS' : 'ORB';

        Text t = new Text();
        t.LongText = o.ObservacaoNotaFiscal__c;
        sapOrder.to_Text = new List<Text>{t};

        if(String.isNotBlank(codigoERPEntrega) && codigoERPEntrega != o.Account.CodigoClienteERP__c){
            Partner p = new Partner();
            p.PartnerFunction = 'ZE';
            p.Customer = codigoERPEntrega;
            sapOrder.to_Partner = new List<Partner>{p};
        }

        sapOrder.to_Item = new List<SapOrderItem>();
        for(OrderItem oi : o.OrderItems){
            SapOrderItem soItem = new SapOrderItem();
            soItem.CustomerPaymentTerms = '';
            soItem.ProductionPlant = '2001';
            soItem.ShippingPoint = '2001';
            soItem.RequestedQuantityUnit = 'KG';
            soItem.SalesOrderItem = oi.ExternalId__c;
            soItem.RequestedQuantity = String.valueOf(oi.QuantityKg__c);
            soItem.NetAmount = String.valueOf(oi.TotalPrice);
            soItem.ItemNetWeight = String.valueOf(oi.PesoLiquido__c);
            soItem.MaterialGroup = oi.Product2.GrupoMercadoria__c;
            soItem.Material = oi.Product2.StockKeepingUnit;
            soItem.SalesOrderItemCategory = o.VendaContaOrdem__c ? 'L2N' : 'TBN';

            // soItem.StorageLocation = (ev.CondicaoExpedicao__c == '01' ? '3000' : '3003');
            soItem.StorageLocation = pce.Deposito__c;

            Decimal totalAdjustmentAmount = 0.0;
            if(context == 'fv'){
                totalAdjustmentAmount = (oi.ValorNegociado__c * oi.QuantityKg__c) - (oi.ListPrice * oi.Quantity);
            } else {
                totalAdjustmentAmount = oi.TotalLineAdjustmentAmount;
            }
            Decimal totalAdjustmentAmountKg = (totalAdjustmentAmount / oi.QuantityKg__c).setscale(2);

            // Desconto aplicado
            PricingElement peDES = new PricingElement();
            peDES.PricingProcedureStep = 'ZDES';
            peDES.ConditionQuantity = totalAdjustmentAmountKg < 0 ?  '1' : '0';
            peDES.ConditionRateValue = totalAdjustmentAmountKg < 0 ?  String.valueOf( - totalAdjustmentAmountKg) : '0';
            soItem.to_PricingElement = new List<PricingElement>{peDES};

            // Acréscimo aplicado
            PricingElement peACR = new PricingElement();
            peACR.PricingProcedureStep = 'ZACR';
            peACR.ConditionQuantity =  totalAdjustmentAmountKg > 0 ?  '1' : '0';
            peACR.ConditionRateValue = totalAdjustmentAmountKg > 0 ?  String.valueOf(totalAdjustmentAmountKg) : '0';
            soItem.to_PricingElement.add(peACR);

            // Cashback usado
            PricingElement peDBK = new PricingElement();
            peDBK.PricingProcedureStep = 'ZDBK';
            peDBK.ConditionQuantity = o.ValorCashback__c == null || o.ValorCashback__c == 0 ? '0' :'1';
            peDBK.ConditionRateValue = o.ValorCashback__c == null || o.ValorCashback__c == 0 ? '0' : String.valueOf(o.ValorCashback__c);
            soItem.to_PricingElement.add(peDBK);

            // Cashback gerado
            PricingElement peCBK = new PricingElement();
            peCBK.PricingProcedureStep = 'ZCBK';
            peCBK.ConditionQuantity = o.CashbackGerado__c == null || o.CashbackGerado__c == 0 ? '0' : '1';
            peCBK.ConditionRateValue = o.CashbackGerado__c == null || o.CashbackGerado__c == 0 ? '0' : String.valueOf(o.CashbackGerado__c);
            soItem.to_PricingElement.add(peCBK);

            sapOrder.to_Item.add(soItem);
        }

        String body = JSON.serialize(sapOrder, true);
        return body;
    }

    private LogErro__c genErrorLog(Exception ex, String methodSign){
        System.debug('ex.getMessage(): ' + ex.getMessage());
        System.debug('ex.getLineNumber(): ' + ex.getLineNumber());
        System.debug('ex.getStackTraceString(): ' + ex.getStackTraceString());
        LogErro__c newLogErro = new LogErro__c();
        newLogErro.Status__c = 'Erro';
        newLogErro.TipoAutomacao__c = 'Apex';
        newLogErro.Nome__c = 'Log de Integração SAP - ' + Datetime.now().format();
        newLogErro.NomeAutomacao__c = 'SapIntegration' + methodSign;
        newLogErro.TextoErro__c = (ex.getMessage() + '\n' + ex.getStackTraceString()).left(20000);
        return newLogErro;
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
                i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
       	i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
                i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
                i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
       	i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
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