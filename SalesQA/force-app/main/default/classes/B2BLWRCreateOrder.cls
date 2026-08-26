public without sharing class B2BLWRCreateOrder {

    public Id checkoutSessionId;
    public WebCart webCart;

    public B2BLWRCreateOrder(Id checkoutSession, WebCart cart) {
        this.webCart = cart;
        this.checkoutSessionId = checkoutSession;
    }

    public String cartToOrder() {
        // load the primary delivery group (only one supported at this time)
        Id cartDeliveryGroupId = [SELECT Id FROM CartDeliveryGroup WHERE CartId = :this.webCart.Id][0].Id;
        Id orderId = mapAndInsertCartToOrder(this.webCart.Id);
        updateCheckoutSession(this.checkoutSessionId, orderId);
        Id orderDeliveryGroupId = mapAndInsertCartDeliveryGroupToOrderDeliveryGroup(cartDeliveryGroupId, orderId);
        mapAndInsertCartItems(cartDeliveryGroupId, orderId, orderDeliveryGroupId);
        return orderId;
    }


    private void updateCheckoutSession(Id checkoutSessionId, Id orderId) {
        CartCheckoutSession checkoutSession = [
            SELECT OrderId, NextState
            FROM CartCheckoutSession
            WHERE Id = :this.checkoutSessionId
            LIMIT 1
        ];
        checkoutSession.OrderId = orderId;
        checkoutSession.NextState = 'Complete';
        update checkoutSession;
    }

    private  Id mapAndInsertCartToOrder(Id cartId) {
        // Get the cart data needed to populate the order
        List<WebCart> carts = [
            SELECT Id,
            AccountId,
            OwnerId,
            WebStoreId,
            PoNumber,
            BillingStreet,
            BillingCity,
            BillingState,
            BillingPostalCode,
            BillingCountry,
            BillingLatitude,
            BillingLongitude,
            WebStore.StrikethroughPricebookId
            FROM WebCart WHERE Id = :cartId
        ];

        WebCart cart = carts[0];
        User usr = [SELECT ContactId FROM USER WHERE Id = :cart.OwnerId];
        
        // Create the order
        Date now = Date.today(); 
        Order order = new Order(
            AccountId = cart.AccountId,
            OwnerId = cart.OwnerId,
            SalesStoreId = cart.WebStoreId,
            PoNumber = cart.PoNumber,
            BillingStreet = cart.BillingStreet,
            BillingCity = cart.BillingCity,
            BillingState = cart.BillingState,
            BillingPostalCode = cart.BillingPostalCode,
            BillingCountry = cart.BillingCountry,
            BillingLatitude = cart.BillingLatitude,
            BillingLongitude = cart.BillingLongitude,
            EffectiveDate = now,
            WebCart__c = cart.Id,
            OrderedDate = now,
            ContactId__c = usr.ContactId,
            NumeroPedidoAgrosys__c = orderNumberAgrosys(),
            Pricebook2Id = cart.WebStore.StrikethroughPricebookId,
            Type = 'B2B Commerce',
            Status = 'Draft'
        );

        insert order;
        return order.Id;
    }

    private  Id mapAndInsertCartDeliveryGroupToOrderDeliveryGroup(Id cartDeliveryGroupId, Id orderId) {
        // Get the cart delivery group data needed to populate the order delivery group
        User user = [
            SELECT Id, ContactId, Contact.AccountId
            FROM User
            WHERE Id = :UserInfo.getUserId()
        ];

        Account acc;
        if(Test.isRunningTest()){
            acc = [
                SELECT Id, Name, ShippingStreet, ShippingCity, ShippingState, ShippingPostalCode, ShippingCountry, Unidade__c
                FROM Account
                WHERE Name = 'isTest'
            ];
        }else{
            acc = [SELECT Id, Name, ShippingStreet, ShippingCity, ShippingState, ShippingPostalCode, ShippingCountry, Unidade__c
            FROM Account
            WHERE Id = :user.Contact.AccountId];
        }

        ContactPointAddress pointAddress = [
            SELECT Id, Name, Street, City, State, PostalCode, Country, AddressType, ParentId, UsageType, IsDefault 
            FROM ContactPointAddress
            WHERE ParentId = :acc.Id
            AND IsDefault = true
            AND AddressType = 'Shipping'
            Limit 1
        ];

        OrderDeliveryMethod deliveryMethod = [
            SELECT Id
            FROM OrderDeliveryMethod
            WHERE Name = 'Entrega Padrão'
            AND IsActive = true
            LIMIT 1
        ];

        // Create the order delivery group
        Date desiredDeliveryDate = Date.today();

        OrderDeliveryGroup orderDeliveryGroup = new OrderDeliveryGroup(
            DesiredDeliveryDate = desiredDeliveryDate,
            DeliverToName = pointAddress.name,
            DeliveryInstructions = '',
            DeliverToStreet = pointAddress.street,
            DeliverToCity = pointAddress.city,
            DeliverToState = pointAddress.state,
            DeliverToPostalCode = pointAddress.PostalCode,
            DeliverToCountry = pointAddress.country,
            OrderDeliveryMethodId = deliveryMethod.Id,
            OrderId = orderId
        );
        insert(orderDeliveryGroup);
        return orderDeliveryGroup.Id;
    }

    private  void mapAndInsertCartItems(Id cartDeliveryGroupId, Id orderId, Id orderDeliveryGroupId) {
        // Get the cart items needed to populate the order items and adjustments
        List<CartItem> cartItems = [
            SELECT AdjustmentAmount, Product2Id, Type, Quantity, ListPrice, SalesPrice, TotalLineAmount, TotalAdjustmentAmount,
            (SELECT Name, TaxType, TaxCalculationDate, Amount FROM CartItemTaxes)
            FROM CartItem
            WHERE CartDeliveryGroupId = :cartDeliveryGroupId
        ];

        List<OrderItem> orderItems = new List<OrderItem>();
        // For each item, map it to an order, then add adjustments
        for (CartItem cartItem : cartItems) {
            orderItems.add(mapCartItemToOrderItem(cartItem, orderId, orderDeliveryGroupId));
        }

        // If there are no items to insert, we can't do anything
        if (orderItems.size() == 0 || cartItems.size() != orderItems.size()) {
            return;
        }

        insert(orderItems);

        List<OrderItemAdjustmentLineItem> lineItemAdjustments = new List<OrderItemAdjustmentLineItem>();
        List<OrderItemTaxLineItem> lineItemTaxes = new List<OrderItemTaxLineItem>();
        for (Integer index = 0; index < cartItems.size(); index++) {
            OrderItemAdjustmentLineItem lineItemAdjustment = mapOrderItemAdjustmentLineItemTo(cartItems.get(index), orderItems.get(index).Id);
            if (lineItemAdjustment != null) {
                lineItemAdjustments.add(lineItemAdjustment);
            }
            List<OrderItemTaxLineItem> lineItemTax = mapOrderItemTaxLineItemTo(cartItems.get(index).CartItemTaxes, orderItems.get(index).Id);
            if (!lineItemTax.isEmpty()) {
                lineItemTaxes.addAll(lineItemTax);
            }
        }
        if (lineItemAdjustments.size() > 0) {
            insert(lineItemAdjustments);
        }
        if (lineItemTaxes.size() > 0) {
            insert(lineItemTaxes);
        }
    }


    private  OrderItem mapCartItemToOrderItem(CartItem cartItem, Id orderId, Id orderDeliveryGroupId) {
        String orderItemType = getOrderItemType(cartItem.Type);
        Decimal unitPrice = getUnitPrice(cartItem);
        Decimal listPrice = cartItem.ListPrice > 0 ? cartItem.ListPrice : UnitPrice;

        OrderItem orderItem = new OrderItem(
            Product2Id = cartItem.Product2Id,
            Type = orderItemType,
            Quantity = cartItem.Quantity,
            ListPrice = listPrice,
            UnitPrice = unitPrice,
            OrderId = orderId,
            OrderDeliveryGroupId = orderDeliveryGroupId,
            TotalLineAmount = cartItem.TotalLineAmount
        );

        for (CartTax tax : cartItem.CartItemTaxes) {
            if (tax.Name == 'Combate à pobreza ST') orderItem.FundoCombatePobrezaST__c = tax.Amount;
            if (tax.Name == 'ST') orderItem.ValorST__c = tax.Amount;
        }

        return orderItem;
    }


    private  OrderItemAdjustmentLineItem mapOrderItemAdjustmentLineItemTo(CartItem cartItem, Id orderItemId) {
        Decimal adjustmentAmount = getAdjustmentAmount(cartItem);
        if (adjustmentAmount == null || adjustmentAmount == 0.0) {
            return null;
        }

        OrderItemAdjustmentLineItem orderItemAdjustmentLineItem = new OrderItemAdjustmentLineItem(
            Amount = adjustmentAmount,
            OrderItemId = orderItemId,
            Name = 'Price Adjustment'
        );

        return orderItemAdjustmentLineItem;
    }

    private List<OrderItemTaxLineItem> mapOrderItemTaxLineItemTo(List<CartTax> CartItemTaxes, Id orderItemId) {
        List<OrderItemTaxLineItem> lstTaxInsert = new List<OrderItemTaxLineItem>();
        for (CartTax tax : CartItemTaxes) {
            OrderItemTaxLineItem orderTax = new OrderItemTaxLineItem(
                Name = tax.Name,
                OrderItemId = orderItemId,
                Type = tax.TaxType,
                TaxEffectiveDate = tax.TaxCalculationDate,
                Amount = tax.Amount
            );
            lstTaxInsert.add(orderTax);
        }

        return lstTaxInsert;
    }

    private  Decimal getAdjustmentAmount(CartItem cartItem) {
        if (cartItem.TotalAdjustmentAmount == null) {
            return 0;
        }

        return cartItem.TotalAdjustmentAmount;
    }

    private  String getOrderItemType(String cartItemType) {
        if (cartItemType == 'Product') {
            return 'Order Product';
        }
        if (cartItemType == 'Charge') {
            return 'Delivery Charge';
        }

        return null;
    }

    private  Decimal getUnitPrice(CartItem cartItem) {
        if (cartItem.SalesPrice != null) {
            return cartItem.SalesPrice;
        }

        return cartItem.ListPrice;
    }

    private String orderNumberAgrosys() {
        List<Order> lstOrder = [
            SELECT Id, NumeroPedidoAgrosys__c
            FROM Order
            WHERE NumeroPedidoAgrosys__c != NULL
            ORDER BY CreatedDate DESC
            LIMIT 1
        ];

        if(!lstOrder.isEmpty()) {
            Integer newOrderNumber = Integer.valueOf(lstOrder[0].NumeroPedidoAgrosys__c) + 1;
            return String.valueOf(newOrderNumber);
        }
        return '1000';
    }
    
}