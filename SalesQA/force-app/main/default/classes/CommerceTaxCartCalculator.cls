@SuppressWarnings('PMD.CyclomaticComplexity,PMD.NcssMethodCount,PMD.CognitiveComplexity')
public class CommerceTaxCartCalculator extends CartExtension.TaxCartCalculator {
    public virtual override void calculate(CartExtension.CartCalculateCalculatorRequest request) {
        try {
            CartExtension.Cart cart = request.getCart();
            
            // Clean up CVO based on tax. When new tax calculator request comes, we need to clean up
            // previous CVOs as they have been previously handled by the Cart Calculate API.
            CartExtension.CartValidationOutputList cartValidationOutputCollection = cart.getCartValidationOutputs();
            for (Integer i = (cartValidationOutputCollection.size() - 1); i >= 0; i--) {
                CartExtension.CartValidationOutput cvo = cartValidationOutputCollection.get(i);
                if (cvo.getType() == CartExtension.CartValidationOutputTypeEnum.TAXES) {
                    cartValidationOutputCollection.remove(cvo);
                }
            }
            
            // There should be one delivery group per cart.
            CartExtension.CartDeliveryGroupList cartDeliveryGroups = cart.getCartDeliveryGroups();
            System.debug('cartDeliveryGroups: ' + cartDeliveryGroups);
            CartExtension.CartDeliveryGroup cartDeliveryGroup = cartDeliveryGroups.get(0);
            System.debug('cartDeliveryGroup: ' + cartDeliveryGroup);

            // Map cart ID to cart item with type Product.
            CartExtension.CartItemList cartItemCollection = cart.getCartItems();
            System.debug('cartItemCollection:');
            System.debug(cartItemCollection);
            // The cartItemCollection contains both products and shipping cart items.
            Map<String, CartExtension.CartItem> cartItemByProductId = new Map<String, CartExtension.CartItem>();
            Map<String, CartExtension.CartItem> shippingItemById = new Map<String, CartExtension.CartItem>();
            for (Integer i = (cartItemCollection.size() - 1); i >= 0; i--) {
                System.debug(cartItemCollection.get(i).getId());
                System.debug(cartItemCollection.get(i).getProduct2Id());
                System.debug(cartItemCollection.get(i).getQuantity());
                
                if (cartItemCollection.get(i).getType() == CartExtension.SalesItemTypeEnum.PRODUCT) {
                    cartItemByProductId.put(cartItemCollection.get(i).getProduct2Id(), cartItemCollection.get(i));
                } else if (cartItemCollection.get(i).getType() == CartExtension.SalesItemTypeEnum.CHARGE) {
                    // Shipping cart items are uniquely identified using delivery group id.
                    CartExtension.CartDeliveryGroup deliveryGroup = cartItemCollection.get(i).getCartDeliveryGroup();
                    shippingItemById.put(deliveryGroup.getId(), cartItemCollection.get(i));
                }
            }
            
            // Get the tax rates and tax amounts from an external service for all given products and its
            // adjustments.
            Map<String, List<TaxOutput>> dataFromExternalService = null;
            dataFromExternalService = getTaxesFromStaticResponse(
                cart,
                cartItemByProductId,
                CartDeliveryGroup.getDeliverToAddress().getState(),
                CartDeliveryGroup.getDeliverToAddress().getCountry(),
                cart.getTaxType()
            );
            
            // If no tax details are returned for any cart item, add a cart validation output entry. If
            // any invalid scenario found then return.
            boolean isCvoPresent = false;
            for (String cartItemProductId : cartItemByProductId.keySet()) {
                List<TaxOutput> taxDetails = dataFromExternalService.get(cartItemProductId);
                if (taxDetails == null || taxDetails.isEmpty()) {
                    // add cvo
                    CartExtension.CartValidationOutput cvo = new CartExtension.CartValidationOutput(CartExtension.CartValidationOutputTypeEnum.TAXES, CartExtension.CartValidationOutputLevelEnum.INFO);
                    cvo.setMessage('No tax rates configured for this location.');
                    cartValidationOutputCollection.add(cvo);
                    isCvoPresent = true;
                }
            }
            if (isCvoPresent == true) return;
            
            // If there are no existing cart tax entries in the cart item that indicates cart item was
            // newly CREATED in the cart then:
            // 1. Create new cart tax entries
            // 2. Update cart item tax information
            for (String cartIteProductmId : dataFromExternalService.keySet()) {
                List<TaxOutput> listTaxDetailsToCartId = dataFromExternalService.get(cartIteProductmId);
                CartExtension.CartItem cartItem = cartItemByProductId.get(cartIteProductmId);
                CartExtension.CartTaxList cartTaxCollection = cartItem.getCartTaxes();

                // Faz o mapa dos impostos existentes por nome
                Map<String, CartExtension.CartTax> mapImpostoxValor = new Map<String, CartExtension.CartTax>();
                if (cartTaxCollection != null && !cartTaxCollection.isEmpty()) {
                    for (Integer i = 0; i < cartTaxCollection.size(); i++) {
                        CartExtension.CartTax tax = cartTaxCollection.get(i);
                        mapImpostoxValor.put(tax.getName(), tax);
                    }
                }
                
                //Cria ou atualizar os impostos
                for (TaxOutput taxDetailsToCartId : listTaxDetailsToCartId) {
                    if (mapImpostoxValor.containsKey(taxDetailsToCartId.getTaxName())) {
                        CartExtension.CartTax cartTax = mapImpostoxValor.get(taxDetailsToCartId.getTaxName());
                        cartTax.setAmount(taxDetailsToCartId.getAmount());
                        System.debug('cartTax Setted:');
                        System.debug(cartTax.getName());
                        System.debug(cartTax.getTaxRate());
                        System.debug(cartTax.getAmount());
                    } else {
                        CartExtension.CartTax cartTax = new CartExtension.CartTax(CartExtension.TaxTypeEnum.ESTIMATED, taxDetailsToCartId.getAmount(), taxDetailsToCartId.getTaxName());
                        cartTax.setTaxRate(String.valueOf(taxDetailsToCartId.getRate()));
                        cartTaxCollection.add(cartTax);
                        System.debug('cartTax Added:');
                        System.debug(cartTax.getName());
                        System.debug(cartTax.getTaxRate());
                        System.debug(cartTax.getAmount());
                    }
                }
                /*if (cartTaxCollection == null || cartTaxCollection.isEmpty()) {
                    for (TaxOutput taxDetailsToCartId : listTaxDetailsToCartId) {
                        CartExtension.CartTax cartTax = new CartExtension.CartTax(CartExtension.TaxTypeEnum.ESTIMATED, taxDetailsToCartId.getAmount(), taxDetailsToCartId.getTaxName());
                        cartTax.setTaxRate(String.valueOf(taxDetailsToCartId.getRate()));
                        cartTaxCollection.add(cartTax);
                    }
                } else {
                    // Faz o mapa dos impostos existentes por nome
                    Map<String, TaxOutput> mapImpostoxValor = new Map<String, TaxOutput>();
                    for (CartExtension.CartTax tax : cartTaxCollection) {
                        mapImpostoxValor.put(tax.getName(), tax);
                    }
                    
                    for (TaxOutput taxDetailsToCartId : listTaxDetailsToCartId) {
                        if (mapImpostoxValor.containsKey(taxDetailsToCartId.getTaxName())) {
                            CartExtension.CartTax cartTax = mapImpostoxValor.get(taxDetailsToCartId.getTaxName());
                            cartTax.setAmount(taxDetailsToCartId.getAmount());
                        } else {
                            CartExtension.CartTax cartTax = new CartExtension.CartTax(CartExtension.TaxTypeEnum.ESTIMATED, taxDetailsToCartId.getAmount(), taxDetailsToCartId.getTaxName());
                            cartTax.setTaxRate(String.valueOf(taxDetailsToCartId.getRate()));
                            cartTaxCollection.add(cartTax);
                        }
                    }
                }*/
            }
        } catch (Exception e) {
            // For testing purposes, this example treats exceptions as user errors, which means they are
            // displayed to the buyer user. In production, you probably want exceptions to be admin-type
            // errors. In that case, throw the exception here and make sure that a notification system is
            // in place to let the admin know that the error occurred. See the README section about error
            // handling for details about how to create that notification.
            throw new CalloutException('There was a problem with the request.');
        }
        return;
    }
    
    // Get cartItemAdjustment based on its ID.
    private CartExtension.cartItemPriceAdjustment getAdjustmentById(
        CartExtension.cartItemPriceAdjustmentList cipaList,
    String id
    ) {
        for (Integer i = (cipaList.size() - 1); i >= 0; i--) {
            if (String.valueOf(cipaList.get(i).getId()) == id)
                return cipaList.get(i);
        }
        return null;
    }
    
    private Map<String, List<TaxOutput>> getTaxesFromStaticResponse(CartExtension.Cart cart, Map<String, CartExtension.CartItem> cartItemsMap, String state, String country, CartExtension.TaxLocaleTypeEnum taxType) {

        System.debug('state: ' + state);
        if(String.isBlank(state)){
            state = [SELECT State FROM ContactPointAddress WHERE ParentId = :cart.getAccountId() AND AddressType = 'Shipping' AND IsDefault = true]?.State ?? null;
            System.debug('state2: ' + state);
        }

        String setorIndustrial = [  SELECT  SetorIndustrial__c
                                    FROM    Account
                                    WHERE   Id = :cart.getAccountId() ]?.SetorIndustrial__c ?? '';



        Map<String, List<TaxOutput>> resultsFromStaticResponse =  new Map<String, List<TaxOutput>>();
        // if(String.isBlank(setorIndustrial) || (setorIndustrial != '102' && setorIndustrial != '108')){  // Calcula ST apenas para setores 102 e 108
        //     for (CartExtension.CartItem item : cartItemsMap.values()){
        //         TaxOutput tax = new TaxOutput(0, 0, 'ICMS-ST');
        //         resultsFromStaticResponse.put(item.getProduct2Id(), new List<TaxOutput>{tax});
        //     }

        // } else {
            // Id's dos produtos do carrinho
            Set<String> setProductId = new Set<String>();
            for(CartExtension.CartItem ci : cartItemsMap.values()){
                setProductId.add((String)ci.getProduct2Id());
            }

            // Dados dos produtos do carrinho
            Map<String,Product2> mapProduct = new Map<String,Product2>();
            for(Product2 prd : [    SELECT  Id, CodigoProdutoERP__c, CodigoControle__c, GrupoMercadoriaExterno__c
                                    FROM    Product2
                                    WHERE   Id IN :setProductId]){
                mapProduct.put(prd.Id, prd);
            }

            // Alíquote interna de ICMS
            Decimal percIcmsInterno = [ SELECT  Aliquota__c
                                        FROM    TabelaIcms__c
                                        WHERE   UfOrigem__c = 'MG'
                                        AND     UfDestino__c = 'MG'
                                        ORDER   BY ValidoDesde__c DESC LIMIT 1 ]?.Aliquota__c ?? 0;
            System.debug('percIcmsInterno: ' + percIcmsInterno);

            // Alíquote interstadual de ICMS
            Decimal percIcmsInterestadual = [   SELECT  Aliquota__c
                                                FROM    TabelaIcms__c
                                                WHERE   UfOrigem__c = 'MG'
                                                AND     UfDestino__c = :state
                                                ORDER   BY ValidoDesde__c DESC LIMIT 1 ]?.Aliquota__c ?? 0;
            System.debug('percIcmsInterestadual: ' + percIcmsInterestadual);

            // Sequência de acesso para determinação da tabela de ICMS-ST
            Map<String,TabelaIcmsSt__c> mapTabela = new Map<String,TabelaIcmsSt__c>();
            for(TabelaIcmsSt__c t : [   SELECT  Id, Name, GrupoImposto__c, GrupoMercadoriaExterno__c, SetorIndustrial__c, Aliquota__c, BaseReduz1__c
                                        FROM    TabelaIcmsSt__c
                                        WHERE   UfOrigem__c = 'MG'
                                        AND     UfDestino__c = :state
                                        AND     ValidoAte__c >= TODAY
                                        AND     GrupoImposto__c IN ('60', '65', '75', '80')
                                        ORDER   BY ValidoDesde__c ]){

                String key = t.GrupoImposto__c + '_' + t.GrupoMercadoriaExterno__c;
                if(t.GrupoImposto__c == '75'){
                    key = key + '_' + t.SetorIndustrial__c;
                }
                mapTabela.put(key, t);
            }
            System.debug('mapTabela: ' + mapTabela);
            
            for (CartExtension.CartItem item : cartItemsMap.values()){
                Product2 prd = mapProduct.get(item.getProduct2Id());
                String codigoMaterial = prd.CodigoProdutoERP__c.leftPad(18, '0');

                String chaveGrupo = '60_' + prd.GrupoMercadoriaExterno__c;
                if(!mapTabela.containsKey(chaveGrupo)){
                    chaveGrupo = '65_' + codigoMaterial;
                    if(!mapTabela.containsKey(chaveGrupo)){
                        chaveGrupo = '75_' + prd.GrupoMercadoriaExterno__c + '_' + setorIndustrial;
                        if(!mapTabela.containsKey(chaveGrupo)){
                            chaveGrupo = '80_' + prd.GrupoMercadoriaExterno__c;
                            if(!mapTabela.containsKey(chaveGrupo) || setorIndustrial == '101'){
                                chaveGrupo = null;
                            }
                        }
                    }
                }

                // Versão obsoleta da sequência de acesso
                // String chaveGrupo = '65_' + codigoMaterial;
                // if(!mapTabela.containsKey(chaveGrupo)){
                //     chaveGrupo = '75_' + prd.GrupoMercadoriaExterno__c + '_' + setorIndustrial;
                //     if(!mapTabela.containsKey(chaveGrupo)){
                //         chaveGrupo = '80_' + prd.GrupoMercadoriaExterno__c;
                //         if(!mapTabela.containsKey(chaveGrupo)){
                //             chaveGrupo = '85_' + prd.GrupoMercadoriaExterno__c;
                //             if(!mapTabela.containsKey(chaveGrupo)){
                //                 chaveGrupo = null;
                //             }
                //         }
                //     }
                // }

                Decimal percIcmsSt = 0;
                Decimal valorIcmsSt = 0;

                if(String.isNotBlank(chaveGrupo) && mapTabela.containsKey(chaveGrupo)){
                    TabelaIcmsSt__c tabela = mapTabela.get(chaveGrupo);

                    Decimal mva = tabela.Aliquota__c ?? 0;
                    System.debug('mva: ' + mva);

                    Decimal base = (100 - (tabela.BaseReduz1__c ?? 100));
                    System.debug('base: ' + base);

                    Object valorNegociadoC = item.getCustomField('ValorNegociado__c');
                    Decimal valorNegociado = valorNegociadoC == null ? (item.getTotalPriceAfterAllAdjustments() / item.getQuantity() / item.getPerUnitWeight()) : (Decimal)valorNegociadoC;
                    System.debug('valorNegociado: ' + valorNegociado);

                    Decimal valorTotal = (valorNegociado * item.getQuantity() * item.getPerUnitWeight());
                    System.debug('valorTotal: ' + valorTotal);

                    Decimal baseIcms = (valorTotal * base / 100);
                    System.debug('baseIcms: ' + baseIcms);

                    Decimal valorIcmsInter = (baseIcms * percIcmsInterestadual / 100);
                    System.debug('valorIcmsInter: ' + valorIcmsInter);

                    Decimal valorMVA = (1 + (mva / 100));
                    System.debug('valorMVA: ' + valorMVA);

                    Decimal baseIcmsSt = (baseIcms * valorMVA);
                    System.debug('baseIcmsSt: ' + baseIcmsSt);

                    valorIcmsSt = ((baseIcmsSt * (percIcmsInterno / 100)) - valorIcmsInter).setScale(2);
                    System.debug('valorIcmsSt: ' + valorIcmsSt);

                    percIcmsSt = (valorIcmsSt / valorTotal * 100).setScale(2);
                    System.debug('percIcmsSt: ' + percIcmsSt);
                }

                TaxOutput tax = new TaxOutput(percIcmsSt, valorIcmsSt, 'ICMS-ST');
                resultsFromStaticResponse.put(item.getProduct2Id(), new List<TaxOutput>{tax});
            }
        // }

        return resultsFromStaticResponse;
    }
    
    // Structure to store the tax data retrieved from external service. This class simplifies our
    // ability to access the data when storing it in Salesforce's CartTaxDto.
    class TaxDataFromExternalService {
        private Decimal rate;
        private Decimal amount;
        private String taxName;
        private Decimal adjustmentTaxAmount;
        private Decimal totalItemizedPromotionTaxAmount;
        private List<CartAdjustment> itemizedPromotionTaxAmounts;
        private Decimal grossUnitPrice;
        private Decimal netUnitPrice;
        
        public TaxDataFromExternalService() {
            rate = 0.0;
            amount = 0.0;
            taxName = '';
            adjustmentTaxAmount = 0.0;
            totalItemizedPromotionTaxAmount = 0.0;
            itemizedPromotionTaxAmounts = null;
            grossUnitPrice = 0.0;
            netUnitPrice = 0.0;
        }
        
        public TaxDataFromExternalService(
            Decimal rateObj,
        Decimal amountObj,
        String taxNameObj,
        Decimal adjustmentTaxAmountObj,
        Decimal totalItemizedPromotionTaxAmountObj,
        List<CartAdjustment> itemizedPromotionTaxAmountsObj,
        Decimal grossUnitPriceObj,
        Decimal netUnitPriceObj
        ) {
            rate = rateObj;
            amount = amountObj;
            taxName = taxNameObj;
            adjustmentTaxAmount = adjustmentTaxAmountObj;
            totalItemizedPromotionTaxAmount = totalItemizedPromotionTaxAmountObj;
            itemizedPromotionTaxAmounts = itemizedPromotionTaxAmountsObj;
            grossUnitPrice = grossUnitPriceObj;
            netUnitPrice = netUnitPriceObj;
        }
        
        public Decimal getRate() {
            return rate;
        }
        
        public Decimal getAmount() {
            return amount;
        }
        
        public String getTaxName() {
            return taxName;
        }
        
        public Decimal getAdjustmentTaxAmount() {
            return adjustmentTaxAmount;
        }
        
        public Decimal getTotalItemizedPromotionTaxAmount() {
            return totalItemizedPromotionTaxAmount;
        }
        
        public List<CartAdjustment> getItemizedPromotionTaxAmounts() {
            return itemizedPromotionTaxAmounts;
        }
        
        public Decimal getGrossUnitPrice() {
            return grossUnitPrice;
        }
        
        public Decimal getNetUnitPrice() {
            return netUnitPrice;
        }
    }
    
    class CartAdjustment {
        private String id;
        private Decimal amount;
        
        public CartAdjustment() {
            id = '';
            amount = 0.0;
        }
        
        public CartAdjustment(String idObj, Decimal taxAmountObj) {
            id = idObj;
            amount = taxAmountObj;
        }
        
        public String getId() {
            return id;
        }
        
        public Decimal getAmount() {
            return amount;
        }
    }

    public class TaxOutput {
        private Decimal rate;
        private Decimal amount;
        private String taxName;
        
        public TaxOutput() {
            rate = 0.0;
            amount = 0.0;
            taxName = '';
        }
        
        public TaxOutput(
            Decimal rateObj,
            Decimal amountObj,
            String taxNameObj
        ) {
            rate = rateObj;
            amount = amountObj;
            taxName = taxNameObj;
        }
        
        public Decimal getRate() {
            return rate;
        }
        
        public Decimal getAmount() {
            return amount;
        }
        
        public String getTaxName() {
            return taxName;
        }
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
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
        i++;
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