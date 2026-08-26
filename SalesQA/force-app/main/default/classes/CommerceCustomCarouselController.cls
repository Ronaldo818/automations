public without sharing class CommerceCustomCarouselController {
    
    @AuraEnabled (cacheable=true)
    public static String getImageUrl(String listContentKeys, String channelName){
        List<String> keys = listContentKeys.split(',', 59);
        DateTime dateNow = DateTime.now();
        String startDate = '2000-01-01T00:00:00.000Z';
        String endDate = dateNow.format('yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
        ConnectApi.ManagedContentVersionCollection obj;
        if(!Test.isRunningTest()) {
            String channelId = [SELECT Id FROM ManagedContentChannel WHERE Name = :channelName LIMIT 1].id;
            obj = ConnectApi.ManagedContent.getContentByContentKeys(channelId, keys, 0, 25, 'pt_BR', 'cms_image', false, startDate, endDate, true);  
        }
        else {
            obj = new ConnectApi.ManagedContentVersionCollection();
            obj.items = new List<ConnectApi.ManagedContentVersion>();
            ConnectApi.ManagedContentVersion item = new ConnectApi.ManagedContentVersion();
            item.title = 'isTest';
            ConnectApi.ManagedContentMediaSourceNodeValue newSourceNodeValue = new ConnectApi.ManagedContentMediaSourceNodeValue();
            newSourceNodeValue.url = 'isTest';
            
            Map<String, ConnectApi.ManagedContentNodeValue> mapManagedNodeValue = new Map<String, ConnectApi.ManagedContentNodeValue>();
            mapManagedNodeValue.put('source', newSourceNodeValue);
            item.contentNodes = mapManagedNodeValue;
            obj.items.add(item);
        }
        
        List<ReturnWrapper> wrapperList = new List<ReturnWrapper>();
       
        for(ConnectApi.ManagedContentVersion versionObj : obj.items){
            ReturnWrapper wrapper = new ReturnWrapper();
            wrapper.title = versionObj.title;
            Map<String, ConnectApi.ManagedContentNodeValue> contentNodesMap = versionObj.contentNodes;
            for(String str : contentNodesMap.keySet()){                
                if(str == 'source'){
                    wrapper.url = ((ConnectApi.ManagedContentMediaSourceNodeValue)contentNodesMap.get(str)).url;
                }		
            }
            wrapperList.add(wrapper);	
        }
        return JSON.serialize(wrapperList);
    }

    public class ReturnWrapper{
        String title {get;set;}
        String url {get;set;}
    }
}