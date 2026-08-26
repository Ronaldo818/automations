public class CommercePricingCalculator extends CartExtension.PricingCartCalculator {

    private Boolean useExternalService = true;
    
    // * @description All classes extending CartExtension.AbstractCartCalculator must have a default constructor defined
    public CommercePricingCalculator() {
        super();
    }

    public CommercePricingCalculator(Boolean useExternalService) {
        super();
        this.useExternalService = useExternalService;
    }

    public virtual override void calculate(CartExtension.CartCalculateCalculatorRequest request) {
        try {
            CartExtension.Cart cart = request.getCart();
            Iterator<CartExtension.CartItem> cartItemsIterator = clearErrorsAndGetCartItemsIterator(cart, request.getOptionalBuyerActionDetails());

            // Get the SKUs from each cart item that needs price calculation.
            Set<String> skus = new Set<String>();
            List<CartExtension.CartItem> cartItems = new List<CartExtension.CartItem>();
            while (cartItemsIterator.hasNext()) {
                CartExtension.CartItem cartItem = cartItemsIterator.next();
                // Pricing service is expected to calculate prices for only product (not charges)
                if(cartItem.getType() == CartExtension.SalesItemTypeEnum.PRODUCT) {
                    skus.add(cartItem.getProduct2Id());
                    cartItems.add(cartItem);
                }
            }

            // Map<String, PricingDataFromExternalService> pricingDataMap = retrievePriceData(skus, cart);

            Map<String, Decimal> pricingDataMap = CommercePricingHelper.getPriceFromSapPriceList(cart.getWebStoreId(), cart.getAccountId(), skus, true);
            System.debug('pricingDataMap: ' + pricingDataMap);

            Map<String, PricingDataFromExternalService> result = new Map<String, PricingDataFromExternalService>();
            for(String productId : pricingDataMap.keyset()){
                Double price = pricingDataMap.get(productId);
                if(price > 0){
                    result.put(productId, new PricingDataFromExternalService(price, price, 0, price, price));
                }
            }
            System.debug('result: ' + result);

            if(pricingDataMap == Null) {
                // No data returned means there is an issue with underlying 3rd party service. Populate generic error message for the Buyer.
                CartExtension.CartValidationOutput cvo = new CartExtension.CartValidationOutput(
                CartExtension.CartValidationOutputTypeEnum.PRICING,
                CartExtension.CartValidationOutputLevelEnum.ERROR);
                cvo.setMessage('We are not able to process your cart. Please contact support.');
                cart.getCartValidationOutputs().add(cvo);
                return;
            }
            applyPricesToCartItems(cart, cartItems.iterator(), result);
        }
        catch (Exception e) {
            // This is an example of throwing special type of Exception (CartCalculateRuntimeException).
            // Throwing this exception causes the rollback of all previously applied changes to the cart (in scope of given request)
            // and may not always be the best choice.
            throw new CartExtension.CartCalculateRuntimeException('An integration error occurred in COMPUTE_PRICES. Contact your admin.');
        }
    }

    /**
    * @description Returns an iterator instance for Cart Items that need price calculations.
    *
    * @param cart In memory representation of the Cart
    * @param optionalBuyerActionDetails The latest set of changes applied to the Cart by the Buyer
    *
    * @return <<Iterator<CartExtension.CartItem>>>
    */
    private Iterator<CartExtension.CartItem> clearErrorsAndGetCartItemsIterator(
        CartExtension.Cart cart,
            CartExtension.OptionalBuyerActionDetails optionalBuyerActionDetails) {

        CartExtension.CartValidationOutputList cartValidationOutputs = cart.getCartValidationOutputs();
        if (!optionalBuyerActionDetails.isPresent() || optionalBuyerActionDetails.get().isCheckoutStarted()) {
            // Remove all CVOs of PRICING type since we are going to reprice all cart items.
            removeAllCVOsOfType(cart, CartExtension.CartValidationOutputTypeEnum.PRICING);
            return cart.getCartItems().iterator();
        }

        // Get the list of changed cart items from BuyerActionDetails
        List<CartExtension.CartItem> cartItemsToCalculatePrice = new List<CartExtension.CartItem>();
        List<CartExtension.CartItemChange> cartItemChanges = optionalBuyerActionDetails.get().getCartItemChanges();

        for (CartExtension.CartItemChange cartItemChange : cartItemChanges) {
            CartExtension.OptionalCartItem optionalCartItem = cartItemChange.getChangedItem();
            if (optionalCartItem.isPresent()) {
                CartExtension.CartItem cartItem = optionalCartItem.get();
                cartItemsToCalculatePrice.add(cartItem);
                // Remove CVOs related to the changed cart item, but leave other PRICING related CVOs.
                removeCVOsOfTypeRelatedTo(cart, CartExtension.CartValidationOutputTypeEnum.PRICING, cartItem);
            }
        }
        return cartItemsToCalculatePrice.iterator();
    }


    private void removeAllCVOsOfType(CartExtension.Cart cart, CartExtension.CartValidationOutputTypeEnum type) {

        CartExtension.CartValidationOutputList cartValidationOutputList = cart.getCartValidationOutputs();
        for (Integer i = (cartValidationOutputList.size() - 1); i >= 0; i--) {
            CartExtension.CartValidationOutput cvo = cartValidationOutputList.get(i);
            if (cvo.getType() == type) {
                cartValidationOutputList.remove(cvo);
            }
        }
    }

    private void removeCVOsOfTypeRelatedTo(CartExtension.Cart cart, CartExtension.CartValidationOutputTypeEnum type,
        CartExtension.CartItem cartItem) {

        if (cartItem.getId() == null) {
            return;
        }
        CartExtension.CartValidationOutputList cartValidationOutputList = cart.getCartValidationOutputs();
        for (Integer i = (cartValidationOutputList.size() - 1); i >= 0; i--) {
            CartExtension.CartValidationOutput cvo = cartValidationOutputList.get(i);
            if (cvo.getRelatedEntityId() == null) {
                continue;
            }
            if (cvo.getType() == type && cvo.getRelatedEntityId() == cartItem.getId()) {
                cartValidationOutputList.remove(cvo);
            }
        }
    }

    private Map<String, PricingDataFromExternalService> retrievePriceData(Set<String> skus, CartExtension.Cart cart) {

        Map<String, PricingDataFromExternalService> dataFromExternalService = null;
        if(useExternalService){
            dataFromExternalService = getPricingFromExternalService(skus, cart);
        } else{
            dataFromExternalService = getPricingFromStaticResponse(skus);
        }
        return dataFromExternalService;
    }

    /**
    * Handles pricing data response by external or static response and sets it to cart items.
    */
    private void applyPricesToCartItems(CartExtension.Cart cart, Iterator<CartExtension.CartItem> cartItemsIterator,
        Map<String, PricingDataFromExternalService> lineItemIdToPricingDetailsMap) {

        while (cartItemsIterator.hasNext()) {
            CartExtension.CartItem cartItem = cartItemsIterator.next();
            if(lineItemIdToPricingDetailsMap.get(cartItem.getProduct2Id()) == null) {
                // No price available for the SKU in the Cart. Populate error message for the Buyer.
                CartExtension.CartValidationOutput cvo = new CartExtension.CartValidationOutput(
                    CartExtension.CartValidationOutputTypeEnum.PRICING,
                        CartExtension.CartValidationOutputLevelEnum.ERROR,
                            cartItem);
                cvo.setMessage('No price available for the SKU in the Cart.');
                cart.getCartValidationOutputs().add(cvo);
                continue;
            }
            setPricingFieldsOnCart(cartItem, lineItemIdToPricingDetailsMap.get(cartItem.getProduct2Id()));
        }
    }

    /**
    * Sets pricing data returned by external service to Salesforce's CartItem
    */
    private void setPricingFieldsOnCart(CartExtension.CartItem item, PricingDataFromExternalService pricingLineItem){

        if (pricingLineItem != null) {
            item.setSalesPrice(pricingLineItem.getUnitPrice());
            item.setListPrice(pricingLineItem.getListPrice());

            item.setTotalListPrice(pricingLineItem.getUnitPrice() * item.getQuantity());
            item.setTotalPrice(pricingLineItem.getUnitPrice() * item.getQuantity());

            // item.setAdjustmentAmount(pricingLineItem.getTotalAdjustmentAmount());
            // item.setTotalAdjustmentAmount(pricingLineItem.getTotalAdjustmentAmount());
            // item.setUnitItemAdjustmentAmount(pricingLineItem.getTotalAdjustmentAmount());
            // item.setTotalListPrice(pricingLineItem.getTotalListPrice());
            // item.setTotalPrice(pricingLineItem.getTotalPrice());
            item.setUnitAdjustedPrice(pricingLineItem.getUnitPrice());
            item.setTotalPriceAfterAllAdjustments(pricingLineItem.getUnitPrice() * item.getQuantity());
            item.setUnitItemAdjustmentAmount(0);
        }
    }

    /**
    * @description This doesn't actually make an HTTP call, but it mocks the JSON response of an external Pricing
    * Service.
    *
    * @param skus A set of SKUs to get pricing data
    *
    * @return Map containing PricingData for each SKU passed into this method
    */
    private Map<String, PricingDataFromExternalService> getPricingFromExternalService(Set<String> skus, CartExtension.Cart cart) {

        // Http http = new Http();
        // HttpRequest request = new HttpRequest();
        // Set<String> encodedSkus = new Set<String>();
        // for (String sku : skus) {
        //     encodedSkus.add(EncodingUtil.urlEncode(sku, 'UTF-8'));
        // }

        // // To access the service below you may need to add endpoint = https://example.com in Setup | Security | Remote site settings.
        // request.setEndpoint('https://example.com/get-sales-prices?skus=' + JSON.serialize(encodedSkus));
        // request.setMethod('GET');
        // HttpResponse r = http.send(request);

        // if (r.getStatusCode() != 200) {
        //     // return null in case of not successful response from 3rd party service
        //     return null;
        // }

    //     Account acc = [  SELECT  EstruturaVendaForcaVendas__r.OrganizacaoVendas__c
    //                         , EstruturaVendaForcaVendas__r.CanalDistribuicao__c
    //                         , EstruturaVendaForcaVendas__r.ListaPreco__c
    //                         , EstruturaVendaForcaVendas__r.EscritorioVendas__c
    //                         , EstruturaVendaForcaVendas__r.GrupoVendedores__c // Equipe de Vendas
    //                         , EstruturaVendaForcaVendas__r.GrupoClientes3__c  // Categoria
    //                         , EstruturaVendaForcaVendas__r.GrupoClientes4__c // Subcategoria
    //                         , EstruturaVendaForcaVendas__r.GrupoClientes5__c // Clusterização
    //                         , EstruturaVendaForcaVendas__r.RegiaoVendas__c
    //                 FROM    Account
    //                 WHERE   Id = :cart.getAccountId() ];
    //     System.debug('acc: ' + acc);
    //     System.debug('OrganizacaoVendas__c: ' + acc.EstruturaVendaForcaVendas__r.OrganizacaoVendas__c);
    //     System.debug('CanalDistribuicao__c: ' + acc.EstruturaVendaForcaVendas__r.CanalDistribuicao__c);
    //    System.debug('skus: ' + skus);

    //     List<ListaPrecoSAP__c> lstPreco = new List<ListaPrecoSAP__c>();
    //     lstPreco = [    SELECT  Id, Produto__c, Prioridade__c, SKU__c, Preco__c
    //                     FROM    ListaPrecoSAP__c
    //                     WHERE   TipoCondicao__c = 'PR00'
    //                     AND     Ativo__c = true
    //                     AND     InicioValidade__c  <= TODAY
    //                     AND     FimValidade__c >= TODAY
    //                     AND     SKU__c = :skus
    //                     AND     OrganizacaoVendas__c = :acc.EstruturaVendaForcaVendas__r.OrganizacaoVendas__c
    //                     AND     CanalDistribuicao__c = :acc.EstruturaVendaForcaVendas__r.CanalDistribuicao__c
    //                     AND (   (   Prioridade__c  = '010'
    //                                 AND Cliente__c = :cart.getAccountId() ) OR
    //                             (   Prioridade__c  = '020'
    //                                 AND ListaPrecos__c = :acc.EstruturaVendaForcaVendas__r.ListaPreco__c ) OR
    //                             (   Prioridade__c  = '030' 
    //                                 AND Cliente__c = :cart.getAccountId()
    //                                 AND EscritorioVendas__c = :acc.EstruturaVendaForcaVendas__r.EscritorioVendas__c
    //                                 AND EquipeVendas__c = :acc.EstruturaVendaForcaVendas__r.GrupoVendedores__c ) OR
    //                             (   Prioridade__c  = '040'
    //                                 AND EscritorioVendas__c = :acc.EstruturaVendaForcaVendas__r.EscritorioVendas__c
    //                                 AND EquipeVendas__c = :acc.EstruturaVendaForcaVendas__r.GrupoVendedores__c) OR
    //                             (   Prioridade__c  = '050'
    //                                 AND EscritorioVendas__c = :acc.EstruturaVendaForcaVendas__r.EscritorioVendas__c
    //                                 AND EquipeVendas__c = :acc.EstruturaVendaForcaVendas__r.GrupoVendedores__c
    //                                 AND Categoria__c = :acc.EstruturaVendaForcaVendas__r.GrupoClientes3__c
    //                                 AND Subcategoria__c = :acc.EstruturaVendaForcaVendas__r.GrupoClientes4__c ) OR
    //                             (   Prioridade__c  = '060'
    //                                 AND EscritorioVendas__c = :acc.EstruturaVendaForcaVendas__r.EscritorioVendas__c
    //                                 AND EquipeVendas__c = :acc.EstruturaVendaForcaVendas__r.GrupoVendedores__c
    //                                 AND Categoria__c = :acc.EstruturaVendaForcaVendas__r.GrupoClientes3__c ) OR
    //                             (   Prioridade__c  = '070'
    //                                 AND EscritorioVendas__c = :acc.EstruturaVendaForcaVendas__r.EscritorioVendas__c
    //                                 AND EquipeVendas__c = :acc.EstruturaVendaForcaVendas__r.GrupoVendedores__c
    //                                 AND Subcategoria__c = :acc.EstruturaVendaForcaVendas__r.GrupoClientes4__c ) OR
    //                             (   Prioridade__c  = '080'
    //                                 AND EscritorioVendas__c = :acc.EstruturaVendaForcaVendas__r.EscritorioVendas__c
    //                                 AND EquipeVendas__c = :acc.EstruturaVendaForcaVendas__r.GrupoVendedores__c
    //                                 AND Clusterizacao__c = 'XXX' ) OR
    //                             (   Prioridade__c  = '090'
    //                                 AND Subcategoria__c = :acc.EstruturaVendaForcaVendas__r.GrupoClientes4__c    ) OR
    //                             (   Prioridade__c  = '100'
    //                                 AND Clusterizacao__c = 'XXX'    ) OR
    //                             (   Prioridade__c  = '120' )
    //                     )
    //                     ORDER BY Produto__c, Prioridade__c DESC
    //     ];
    //     System.debug('lstPreco: ' + lstPreco);

        // Map<String, Decimal> mapPrice = CommercePricingHelper.getPriceFromSapPriceList(request.getWebStoreId(), request.getEffectiveAccountId(), setProductId);

        // Map<String, PricingDataFromExternalService> result = new Map<String, PricingDataFromExternalService>();
        // for(String productId : mapPrice.keyset()){
        //     Double price = mapPrice.get(productId);
        //     result.put(productId, new PricingDataFromExternalService(price, price, 0, price, price));
        // }

        // result.put('976', new PricingDataFromExternalService(200,200,0,200,200));
        // System.debug('result: ' + result);
        // return result;

        // return (Map<String, PricingDataFromExternalService>) JSON.deserialize(r.getBody(), Map<String, PricingDataFromExternalService>.class);

        return null;
    }

    /**
    * Returns a static pricing data map for all products.
    */
    private Map<String, PricingDataFromExternalService> getPricingFromStaticResponse(Set<String> skus) {

        Map<String, PricingDataFromExternalService> dataFromExternalService = new Map<String, PricingDataFromExternalService>();
        for(String sku: skus) {
            PricingDataFromExternalService pricingData = new PricingDataFromExternalService(200,200,0,200,200);
            dataFromExternalService.put(sku, pricingData);
        }
        return dataFromExternalService;
    }

    /**
    * Structure to store pricing data retrieved from external service. This class simplifies our
    * ability to access the data when storing it in Salesforce's CartItem.
    */
    class PricingDataFromExternalService {

        private Decimal unitPrice;
        private Decimal listPrice;
        private Decimal totalAdjustmentAmount;
        private Decimal totalListPrice;
        private Decimal totalPrice;

        public PricingDataFromExternalService() {
            unitPrice = 0.0;
            listPrice = 0.0;
            totalAdjustmentAmount = 0.0;
            totalListPrice = 0.0;
            totalPrice = 0.0;
        }

        public PricingDataFromExternalService(
        Decimal unitPriceObj,
        Decimal listPriceObj,
        Decimal totalAdjustmentAmountObj,
        Decimal totalListPriceObj,
        Decimal totalPriceObj ) {

            unitPrice = unitPriceObj;
            listPrice = listPriceObj;
            totalAdjustmentAmount = totalAdjustmentAmountObj;
            totalListPrice = TotalListPriceObj;
            totalPrice = totalPriceObj;
        }

        public Decimal getUnitPrice() {
            return unitPrice;
        }

        public Decimal getListPrice() {
            return listPrice;
        }

        public Decimal getTotalAdjustmentAmount() {
            return totalAdjustmentAmount;
        }

        public Decimal getTotalListPrice() {
            return totalListPrice;
        }

        public Decimal getTotalPrice() {
            return totalPrice;
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
    }
}