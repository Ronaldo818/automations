@IsTest
public class CommerceCustomCarouselControllerTest {
	
    @IsTest
    static void getImageUrlSuccess01() {
        String listContentKeys = '000000,1111';
        String channelName = 'isTest';
        String url = CommerceCustomCarouselController.getImageUrl(listContentKeys, channelName);
    }
}