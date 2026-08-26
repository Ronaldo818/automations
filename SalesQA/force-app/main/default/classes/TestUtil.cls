@isTest
global class TestUtil {

    
    public String paymentProvider;
    Id mainSimpleProductId;
    Id mainStoreId;
    Id standardPricebookId;
    Id mainContactId;
    Id mainAccountId;
    Id mainBlogId;
    String context = 'b2b'; // b2b, fv    

    public TestUtil(String c ){
        context = c;
        init();
    }

    public TestUtil(){
        init();
    }

    private void init(){
        if(Test.isRunningTest()){
            standardPricebookId = Test.getStandardPricebookId();
            Pricebook2 pb = new Pricebook2();
            pb.Id = Test.getStandardPricebookId();
            pb.CurrencyIsoCode = 'USD';
            update pb;
        }
    }


    private Id getMainContactId(){
        if(mainContactId == null){
            genAccountContactData();
            return mainContactId;
        }
        else{
            return mainContactId;
        }
    }


    public Id getMainAccountId(){
        if(mainAccountId == null){
            genAccountContactData();
            return mainAccountId;
        }
        else{
            return mainAccountId;
        }
    }

    public Id getMainStoreId(){
        if(mainStoreId == null){
            genStoreData();
            return mainStoreId;
        }
        else{
            return mainStoreId;
        }
    }

    private Id getMainSimpleProductId(){
        if(mainSimpleProductId == null){
            genProductsData();
            return mainSimpleProductId;
        }
        else{
            return mainSimpleProductId;
        }
    }

    String shippingProductId;
    private Id getShippingProductId(){
        if(shippingProductId == null){
            genOrderData();
            return shippingProductId;
        }
        else{
            return shippingProductId;
        }
    }

    public void genAllData(){
        genOrderData();
        genProductsData();
        genStoreData();
        genMasterOrderData();
    }

    public void genProductsData(){
        if(Test.isRunningTest() && mainSimpleProductId == null){
            // Simple
            Product2 ps = new Product2();
            ps.Name =  'TESTE LINGUICA FRANGO CONG FINA PCT 2,5KG';
            ps.CodigoProdutoERP__c =  'TG11';
            ps.NumeroAntigo__c =  '';
            ps.ProductCode =  'TG11';
            ps.StockKeepingUnit =  'TEST-TG11';
            // ps.Unidade__c =  '1';
            ps.UnidadeMedida__c =  'KG';
            ps.PesoUnitario__c =  20.0;
            ps.IsActive =  true;
            ps.GrupoMercadoria__c =  '';
            ps.GrupoMercadoriaExterno__c =  '';
            ps.PesoEmbalagem__c =  0;
            ps.Volume__c =  0.0;
            ps.CodigoControle__c =  '';
            ps.ProducaoInterna__c =  false;
            ps.CanalDistribuicao__c =  '';
            ps.UmEan__c =  '';
            ps.EAN__c =  '';
            ps.CodigoFamilia__c =  '55';
            ps.CodigoLinhaProduto__c =  '26';
            ps.CodigoOrigemProduto__c =  '9';
            ps.CodigoUnidadeMedidaVenda__c =  '';
            ps.CodigoUnidadeMedidaBasica__c =  '';
            ps.ShippingWeight = 20;
            insert ps;

            mainSimpleProductId = ps.id;

            PricebookEntry pbe_ps= new PricebookEntry();
            pbe_ps.Product2Id = ps.Id;
            pbe_ps.CurrencyIsoCode = 'USD';
            pbe_ps.Pricebook2Id = standardPricebookId;
            pbe_ps.UnitPrice = 80.0;
            pbe_ps.IsActive = true;
            insert pbe_ps;

            UnidadeMedida__c um = new UnidadeMedida__c();
            um.CodigoUnidadeMedida__c = 'KG';
            um.Descricao__c = 'QUILOGRAMA';
            um.Name = 'QUILOGRAMA';
            um.TipoMedida__c = false;
            insert um;

            UnidadeProduto__c up = new UnidadeProduto__c();
            up.Key__c = '1-1234';
            up.Product__c = ps.Id;
            up.SaldoEstoque__c = 100;
            up.Unidade__c = 1;
            up.UnidadeFaturamento__c = um.Id;
            up.UnidadePedido__c = um.Id;
            insert up;

            Estoque__c psStock = new Estoque__c();
            psStock.ExternalId__c = 'TG11-2001-1000-';
            psStock.Produto__c = ps.Id;
            psStock.Cliente__c = null;
            psStock.Centro__c = '2001';
            psStock.Deposito__c = '1000';
            psStock.Lote__c = '';
            psStock.TipoEstoque__c ='01';
            psStock.UnidadeMedida__c ='UN';
            psStock.TipoProduto__c ='HAWA';
            psStock.DataFabricacao__c =null;
            psStock.DataValidade__c =null;
            psStock.EstoqueLivreMmbe__c = 36125.108;
            psStock.QuantidadePendenteVbbe__c = 115;
            psStock.QtdeEstoque__c = 36010.108;
            insert psStock;
        }
    }


    public void genAccountContactData(){
        if(Test.isRunningTest()){
            Cidade__c city = new Cidade__c();
            city.CodigoCidadeIBGE__c = '3556206';
            city.Name = 'Teste';
            insert city;
            Account acc = new Account();
            acc.Name = 'Test ' + context ;
            acc.CodigoClienteERP__c =  '1000020250X'; // Código SAP do Cliente
            acc.Name =  'Empresa Exemplo Ltda';
            acc.NomeFantasia__c =  'Exemplo';
            acc.Ativo__c =  true;
            acc.CodigoIbgeComercial__c =  '3556206';
            acc.CidadeComercial__c =  city.Id;
            acc.BairroComercial__c =  'Centro';
            acc.NumeroComercial__c =  '100';
            acc.CEPComercial__c =  '01001-000';
            acc.RuaComercial__c =  'Av. Paulista';
            acc.CodigoIbgeCobranca__c =  '3556206';
            acc.CidadeCobranca__c =  city.Id;
            acc.BairroCobranca__c =  'Bela Vista';
            acc.NumeroCobranca__c =  '200';
            acc.CEP__c =  '01002-000';
            acc.RuaCobranca__c =  'Rua Augusta';
            acc.CodigoIbgeEntrega__c =  '3556206';
            acc.CidadeEntrega__c =  city.Id;
            acc.BairroEntrega__c =  'Jardins';
            acc.NumeroEntrega__c =  '300';
            acc.CEPEntrega__c =  '01401-000';
            acc.RuaEntrega__c =  'Alameda Santos';
            acc.Email__c =  'contato@exemplo.com.br';
            acc.EmailComprasComercial__c =  'compras@exemplo.com.br';
            acc.EmailBoleto__c =  'financeiro@exemplo.com.br';
            acc.EmailXml__c =  'xml@exemplo.com.br';
            acc.TelefoneCelular__c =  '(11) 91234-5678';
            acc.TelefoneComercial__c =  '(11) 4002-8922';
            acc.TelefoneFinanceiro__c =  '(11) 4002-1122';
            acc.CPFInput__c =  '12345678910';
            acc.CNPJInput__c =  '';
            acc.InscricaoEstadual__c =  '123456789';
            acc.AplicaST__c =  false;
            acc.AnoFundacao__c =  1999;
            acc.TipoParceiroNegocios__c =  'Distribuidor';
            acc.Unidade__c = 1;
            insert acc;

            Contact c = new Contact();
            c.FirstName = 'Test';
            Set<String> contactFields = Schema.SObjectType.Contact.fields.getMap().keySet();
            if(contactFields.contains('inputlastname__c')) {
                c.put('inputlastname__c', 'Test ' + context);
            } 
            c.LastName = 'Test ' + context;
            c.Phone = '+56988888888';
            c.Email = 'test.' + context + '@test.com';
            c.AccountId = acc.Id;
            c.CPFInput__c = '079.544.910-05'; // cpf aleatório
            insert c;

            mainContactId = c.id;
            mainAccountId = acc.Id;

            if(context == 'b2b'){
                ContactPointAddress cpa = new ContactPointAddress();
                cpa.Name = 'Test';
                cpa.Street = 'Test Test';
                cpa.City = 'Test';
                cpa.State = 'Test';
                cpa.PostalCode = 'Test';
                cpa.IsDefault = true;
                cpa.AddressFirstName = c.FirstName;
                cpa.AddressLastName = c.LastName;
                cpa.Country = 'CL';
                cpa.AddressType = 'Shipping';
                cpa.ParentId = acc.Id;
                insert cpa;

                EstruturaVenda__c ev = new EstruturaVenda__c();
                ev.ExternalId__c = acc.ExternalId__c + '_2000_TESTE_6000_100_TESTE' ;// "10000202501_SP_ONLINE", // Montar chave única para a estrutura: Código do Cliente + '_' + Organização de Venda + Canal + Escritório + Grupo de Vendedor + Setor
                ev.CondicaoPagamento__c = 'TESTE';// "30 dias",
                ev.EscritorioVendas__c = '6000';// "SP",
                ev.GrupoVendedores__c = 'TESTE';// "Grupo A",
                ev.SetorAtividade__c = 'TESTE';// "Tecnologia",
                ev.CanalDistribuicao__c = 'TESTE';// "Online",
                ev.OrganizacaoVendas__c = '2000';// "Org 1",
                ev.ListaPreco__c = 'TESTE';// "LP01",
                ev.RegiaoVendas__c = 'TESTE';// "Sudeste",
                ev.GrupoClientes__c = 'TESTE';// "GC1",
                ev.GrupoClientes1__c = 'TESTE';// "GC2",
                ev.GrupoClientes2__c = 'TESTE';// "GC3",
                ev.GrupoClientes3__c = 'TESTE';// "GC4",
                ev.GrupoClientes4__c = 'TESTE';// "GC5",
                ev.GrupoClientes5__c = 'TESTE';// "GC6"
                ev.AccountId__c = acc.Id;
                insert ev;

                Cashback__c cb = new Cashback__c();
                cb.Cliente__c = acc.Id;
                insert cb;
            
                User newUser = new User(
                    Alias 				= 'isTest',
                    username 			=  c.Email,
                    EmailEncodingKey 	= 'UTF-8',
                    LanguageLocaleKey	= 'pt_BR',
                    LocaleSidKey 		= 'pt_BR',
                    TimeZoneSidKey 		= 'America/Sao_Paulo',
                    ProfileId 			= [SELECT Id FROM Profile WHERE Name = 'Commerce B2B' LIMIT 1][0].Id,
                    lastname 			= c.LastName,
                    email 				= c.email,
                    phone 				= c.Phone,
                    ContactId 			= c.Id
                );
                insert newUser;

                //List<PermissionSetAssignment> permissionsAssToInsert = new List<PermissionSetAssignment>();
                //for(PermissionSet ps : [SELECT Id FROM PermissionSet WHERE Name IN ('B2BBuyer', 'B2BBuyer2', 'B2BBuyerGuest', 'B2BCookieConsentGuestUser')]){
                //    PermissionSetAssignment psa = new PermissionSetAssignment();
                //    psa.AssigneeId = newUser.Id;
                //    psa.PermissionSetId = ps.Id;
                //    permissionsAssToInsert.add(psa);
                //}
                //insert permissionsAssToInsert;
            }
            else{
                EstruturaVenda__c ev = new EstruturaVenda__c();
                ev.ExternalId__c = acc.ExternalId__c + '_2000_TESTE_2000_100_TESTE' ;// "10000202501_SP_ONLINE", // Montar chave única para a estrutura: Código do Cliente + '_' + Organização de Venda + Canal + Escritório + Grupo de Vendedor + Setor
                ev.CondicaoPagamento__c = 'TESTE';// "30 dias",
                ev.EscritorioVendas__c = '2000';// "SP",
                ev.GrupoVendedores__c = 'TESTE';// "Grupo A",
                ev.SetorAtividade__c = 'TESTE';// "Tecnologia",
                ev.CanalDistribuicao__c = 'TESTE';// "Online",
                ev.OrganizacaoVendas__c = '2000';// "Org 1",
                ev.ListaPreco__c = 'TESTE';// "LP01",
                ev.RegiaoVendas__c = 'TESTE';// "Sudeste",
                ev.GrupoClientes__c = 'TESTE';// "GC1",
                ev.GrupoClientes1__c = 'TESTE';// "GC2",
                ev.GrupoClientes2__c = 'TESTE';// "GC3",
                ev.GrupoClientes3__c = 'TESTE';// "GC4",
                ev.GrupoClientes4__c = 'TESTE';// "GC5",
                ev.GrupoClientes5__c = 'TESTE';// "GC6"
                ev.AccountId__c = acc.Id;
                insert ev;                
            }
        }
    }

    public void genOrderData(){
        if(Test.isRunningTest()){

            Payment p = new Payment();
            p.CurrencyIsoCode = 	'USD';
            //p.PaymentGroupId = 	9zxFS00000001CDYAY;
            p.AccountId = 	getMainAccountId();
            p.Date = 	Datetime.now();
            p.Amount = 	20300;
            p.Status = 	'Processed';
            p.Type = 	'Sale';
            p.ProcessingMode = 	'External';
            p.GatewayRefNumber = 	'Test';
            p.GatewayResultCode = 	'0';
            p.SfResultCode = 	'Success';
            p.EffectiveDate = 	Datetime.now();
            p.GatewayResultCodeDescription = 	'AUTHORIZED';
            p.GatewayRefDetails = 	'{"card_number":"6623”, payment_type_code":"SI”, installments_number":3,"response_code":0,"authorization_code":"1213"}';
            insert p;


            WebCart cart = new WebCart();
            cart.Name = 	'Cart';
            cart.WebStoreId = getMainStoreId();
            cart.Type = 'Cart';
        	cart.Status = 'Active';
            //cart.CurrencyIsoCode = 	'USD';
            cart.AccountId = 	getMainAccountId();
            //cart.PaymentMethodId = 	8Z7FS00000003qj0AA;
            cart.Status = 	'Closed';
            cart.Type = 	'Cart';
            cart.IsSecondary = 	false;
            cart.IsRepricingNeeded = 	false;
            cart.PoNumber = 	'Test';
            cart.TaxType = 	'Net';
            cart.ValorCashback__c = 0;
            insert cart;

            Order o = new Order();

            o.AccountId = 	getMainAccountId();
            o.IsReductionOrder = 	false;
            o.Status = 	'Draft';
            o.PoNumber = 	'test';
            //o.OrderReferenceNumber = 	'IPB2T-SN3ZZ-D7WDZ-WV26L';
            o.CurrencyIsoCode = 	'USD';
            o.TaxLocaleType = 	'Net';
            //o.Cart__c = 	cart.Id;
            o.EffectiveDate = Date.today();
            //o.Pricebook2Id = standardPricebookId;
            o.BillingPhoneNumber = '+576015801845';
            insert o;

            PaymentGroup pg = new PaymentGroup();
            pg.SourceObjectId = o.Id;
            pg.CurrencyIsoCode = 'USD';
            insert pg;

            p.PaymentGroupId = pg.Id;
            update p;

            cart.PaymentGroupId = pg.Id;
            update cart;

            Product2 shippingCharge = new Product2();
            shippingCharge.Name = 'Shipping Charge';
            shippingCharge.IsActive = true;
            insert shippingCharge;

            shippingProductId = shippingCharge.Id;

            PricebookEntry pbe_shippingCharge= new PricebookEntry();
            pbe_shippingCharge.Product2Id = shippingCharge.Id;
            pbe_shippingCharge.CurrencyIsoCode = 'USD';
            pbe_shippingCharge.Pricebook2Id = standardPricebookId;
            pbe_shippingCharge.UnitPrice = 10.0;
            pbe_shippingCharge.IsActive = true;
            insert pbe_shippingCharge;

            OrderDeliveryMethod odm = new OrderDeliveryMethod();
            odm.Name = 	'Alas Standard';
            odm.CurrencyIsoCode = 	'USD';
            odm.ProductId = 	shippingCharge.Id;
            odm.Carrier = 	'Alas';
            odm.ClassOfService = 	'Standard';
            odm.IsActive = 	true;
            insert odm;

            CartDeliveryGroup  cdg = new CartDeliveryGroup();
            cdg.DeliverToFirstName = null;
            cdg.DeliverToLastName = null;
            cdg.DeliveryMethodId = odm.Id;
            cdg.DeliverToStreet = '1355 West 35th Street';
            cdg.DeliverToState = 'Test';
            cdg.DeliverToCity = 'West Hollywood';
            cdg.DeliverToCountry = 'USA';
            cdg.DeliverToPostalCode = '10012';
            cdg.ShipToPhoneNumber = null;
            cdg.CartId = cart.Id;
            cdg.Name = 'Test Delivery Group';
            insert cdg;


            OrderDeliveryGroup odg = new OrderDeliveryGroup();
            odg.OrderId = 	o.Id;
            odg.OrderDeliveryMethodId = 	odm.Id;
            odg.DeliverToName = 	'Test Test';
            odg.DeliverToStreet = 	'Teste Teste';
            odg.DeliverToCity = 	'Teste';
            odg.DeliverToState = 	'AP';
            odg.DeliverToCountry = 	'CL';
            odg.DeliverToPostalCode = '112041';
            odg.IsGift = 	false;
            odg.EmailAddress = 'test@test.com';
            insert odg;


            OrderItem oi01 = new OrderItem();
            oi01.Product2Id = 	getMainSimpleProductId();
            //oi01.PricebookEntryId = getMainSimplePBEId();
            oi01.OrderId = 	o.Id;
            oi01.Quantity = 10;
            oi01.UnitPrice = 1000;
            oi01.ListPrice = 1000;
            oi01.Description = 'Test Product';
            oi01.TotalLineAmount = 10000;
            oi01.OrderDeliveryGroupId = odg.Id;
            oi01.Type = 'Order Product';
            oi01.LineNumber = 2;
            oi01.listPrice = 100.0;
            insert oi01;

            CartItem ci01 = new CartItem();
            ci01.Name = 'Test Item';
            ci01.Product2Id = 	getMainSimpleProductId();
            ci01.CartId = 	cart.Id;
            ci01.Quantity = 10;
            ci01.ListPrice = 1000;
            ci01.TotalLineAmount = 10000;
            ci01.UnitAdjustedPrice = 1000;
            ci01.TotalPriceAfterAllAdjustments = 10000;
            ci01.TotalPrice = 10000;
            ci01.CartDeliveryGroupId = cdg.Id;
            insert ci01;

            o.Status = 'Activated';
            update o;
        }
    }

    public void genStoreData(){
        if(Test.isRunningTest() && mainStoreId == null){
            String extId;
            String name;
            if(context == 'b2b'){
                extId = 'B2B';
                name = 'Avivar LWR';
            }
            else{
                extId = 'FV';
                name = 'Força de Vendas Avivar	';
            }
            WebStore newStore = new WebStore(
                Name = name , Type = 'B2B', ExternalId__c = extId, OptionsSplitShipmentEnabled = false);
            insert newStore;
            mainStoreId = newStore.id;

            WishList wl = new WishList();
            wl.WebStoreId = mainStoreId;
            wl.Name = 'Test Wishlist';
            wl.AccountId = getMainAccountId();
            insert wl;

            WishlistItem wlItem =   new WishListItem();
            wlItem.WishlistId = wl.id;
            wlItem.Product2Id = getMainSimpleProductId();
            wlItem.Name = 'Test';
            insert wlItem;

            ProductCatalog pc = new ProductCatalog();
            pc.Name = 'Test Catalog';
            insert pc;
        
            WebStoreCatalog wsc = new WebStoreCatalog();
            wsc.ProductCatalogId = pc.Id;
            wsc.SalesStoreId = newStore.Id;
            insert wsc;
        
            ProductCategory pca = new ProductCategory(); 
            pca.name = 'Test';
            pca.CatalogId = pc.id;
            pca.Description = 'Test category';
            insert pca;

            ProductCategoryProduct pcpv = new ProductCategoryProduct();
            pcpv.ProductCategoryId = pca.Id;
            pcpv.IsPrimaryCategory = true;
            pcpv.ProductId = getMainSimpleProductId();
            insert pcpv;

            // cria carrinho atual

            WebCart cart = new WebCart();
            cart.Name = 	'Cart';
            cart.WebStoreId = getMainStoreId();
            cart.Type = 'Cart';
        	cart.Status = 'Active';
            //cart.CurrencyIsoCode = 	'USD';
            cart.AccountId = 	getMainAccountId();
            //cart.PaymentMethodId = 	8Z7FS00000003qj0AA;
            cart.Type = 	'Cart';
            cart.IsSecondary = 	false;
            cart.IsRepricingNeeded = 	false;
            cart.PoNumber = 	'Test';
            cart.TaxType = 	'Net';
            cart.ValorCashback__c = 0;
            insert cart;

            OrderDeliveryMethod odm = new OrderDeliveryMethod();
            odm.Name = 	'Test';
            odm.CurrencyIsoCode = 	'USD';
            odm.ProductId = getShippingProductId();
            odm.Carrier = 	'Alas';
            odm.ClassOfService = 	'Standard';
            odm.IsActive = 	true;
            insert odm;

            
            CartDeliveryGroup  cdg = new CartDeliveryGroup();
            cdg.DeliverToFirstName = null;
            cdg.DeliverToLastName = null;
            cdg.DeliveryMethodId = odm.Id;
            cdg.DeliverToStreet = '1355 West 35th Street';
            cdg.DeliverToState = 'Test';
            cdg.DeliverToCity = 'West Hollywood';
            cdg.DeliverToCountry = 'USA';
            cdg.DeliverToPostalCode = '10012';
            cdg.ShipToPhoneNumber = null;
            cdg.CartId = cart.Id;
            cdg.Name = 'Test Delivery Group';
            insert cdg;

            CartItem ci01 = new CartItem();
            ci01.Name = 'Test Item';
            ci01.Product2Id = 	getMainSimpleProductId();
            ci01.CartId = 	cart.Id;
            ci01.Quantity = 10;
            ci01.ListPrice = 1000;
            ci01.TotalLineAmount = 10000;
            ci01.UnitAdjustedPrice = 1000;
            ci01.TotalPriceAfterAllAdjustments = 10000;
            ci01.TotalPrice = 10000;
            ci01.CartDeliveryGroupId = cdg.Id;
            insert ci01;
        }
    }

    public void genMasterOrderData(){
        MasterOrder__c mo = new MasterOrder__c();
        mo.AccountId__c = getMainAccountId();
        mo.Comment__c = 'Test';
        insert mo;
        
        MasterOrderItem__c moi = new MasterOrderItem__c();
        moi.ProductId__c = getMainSimpleProductId();
        moi.Quantity__c = 9;
        moi.ListPrice__c = 100.0;
        moi.NegotiatedUnitPrice__c = 99.0;
        moi.MasterOrderId__c = mo.Id;
        insert moi;
    }
}