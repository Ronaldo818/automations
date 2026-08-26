public without sharing class CartItemHandler {
    private List<CartItem> oldList;
    private List<CartItem> newList;
    private Map<Id,CartItem> oldMap;
    private Map<Id,CartItem> newMap;
    public class quantityException extends Exception {}

    public CartItemHandler(List<CartItem> oldList, List<CartItem> newList, Map<Id,CartItem> oldMap, Map<Id,CartItem> newMap){
        this.oldList = oldList;
        this.newList = newList;
        this.oldMap = oldMap;
        this.newMap = newMap;
    }

    public void verificarMultiploQuantidade() {
        List<String> listCartId = new List<String>();
        List<String> listProductId = new List<String>();
        try{
            for(CartItem cartItem : newList) {
                listCartId.add(cartItem.CartId);
                listProductId.add(cartItem.Product2Id);
            }

            Map<String, Decimal> mapCartIdxUnidade = new Map<String, Decimal>();
            for (WebCart cart : [SELECT Id, TYPEOF Account WHEN Account THEN Unidade__c END FROM WebCart WHERE Id IN :listCartId]){
                Account acc = cart.Account;
                mapCartIdxUnidade.put(cart.Id, acc.Unidade__c);
            }

            Map<String, Boolean> mapProductIdxTipoMedida = new Map<String, Boolean>();
            Map<String, Decimal> mapProductIdxPesoUnitario = new Map<String, Decimal>();
            for (UnidadeProduto__c unidade : [SELECT Product__c, Unidade__c, UnidadePedido__r.TipoMedida__c, Product__r.ShippingWeight, Product__r.PesoUnitario__c FROM UnidadeProduto__c WHERE Unidade__c IN :mapCartIdxUnidade.values() AND Product__c IN :listProductId]){
                mapProductIdxTipoMedida.put(unidade.Product__c + String.valueOf(unidade.Unidade__c), unidade.UnidadePedido__r.TipoMedida__c);
                Decimal weight = unidade.Product__r.ShippingWeight == null || unidade.Product__r.ShippingWeight == 0 ? unidade.Product__r.PesoUnitario__c : unidade.Product__r.ShippingWeight;
                mapProductIdxPesoUnitario.put(unidade.Product__c, weight);
            }

            for(CartItem cartItem : newList) {
                String key = cartItem.Product2Id + String.valueOf(mapCartIdxUnidade.get(cartItem.CartId));
                if(mapProductIdxTipoMedida.containsKey(key)){
                    if (!mapProductIdxTipoMedida.get(key)) {
                        Decimal pesoUnitario = mapProductIdxPesoUnitario.get(cartItem.Product2Id);
                        if (Math.mod((Integer)cartItem.Quantity, (Integer)pesoUnitario) != 0) {
                            throw new quantityException('A quantidade desse produto deve ser múltiplo de ' + String.valueOf(pesoUnitario));
                        }
                    }
                }
            }
        }
        catch(Exception ex){
           System.debug('CartItemHandler: ' +  ex.getMessage()); 
        }
    }
}