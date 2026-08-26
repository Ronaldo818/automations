public class BuyerAccountHandler {
	
    private List<BuyerAccount> oldList;
    private List<BuyerAccount> newList;
    private Map<Id,BuyerAccount> oldMap;
    private Map<Id,BuyerAccount> newMap;

    public BuyerAccountHandler(List<BuyerAccount> oldList, List<BuyerAccount> newList, Map<Id, BuyerAccount> oldMap, Map<Id, BuyerAccount> newMap){
        this.oldList = oldList;
        this.newList = newList;
        this.oldMap = oldMap;
        this.newMap = newMap;
    }
    
    public void validatePendingCommerceRecords() {
        List<BuyerAccount> lstBuyerAccPendingEcommerceRecords = new List<BuyerAccount>();
        for(BuyerAccount buyerAcc : newList) {
            if(buyerAcc.IsActive && buyerAcc.PendingEcommerceRecords__c) {
                lstBuyerAccPendingEcommerceRecords.add(buyerAcc);
            }
        }

        if(!lstBuyerAccPendingEcommerceRecords.isEmpty() && ScheduleCreateCommerceRecords.validateAsyncApexJob()) {
            ScheduleCreateCommerceRecords.startSchedule();
        }
    }

}