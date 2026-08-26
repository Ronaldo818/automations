public without sharing class MasterOrderApprovalController {

    @AuraEnabled
    public static Map<String,Object> getinitData(Id recordId) {

        MasterOrder__c mo = [   SELECT  Id, RepresentativeId__c, ItensContraProposta__c, Status__c
                                FROM    MasterOrder__c
                                WHERE   Id = :recordId ];

        Boolean enableApproval = (  mo.RepresentativeId__c == UserInfo.getUserId() &&
                                    mo.Status__c == 'Counterproposal' &&
                                    mo.ItensContraProposta__c > 0   );

        Map<String,Object> mapResult = new Map<String,Object>();
        mapResult.put('enableApproval', enableApproval);
        return mapResult;
    }

    @AuraEnabled
    public static void doApproval(String recordId, String status) {
        MasterOrder__c mo = [SELECT Id, Status__c FROM MasterOrder__c WHERE Id = :recordId];
        mo.Status__c = status;
        update mo;
    }

}