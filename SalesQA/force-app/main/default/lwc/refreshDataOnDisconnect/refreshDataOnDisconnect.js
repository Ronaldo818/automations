import { LightningElement, api } from 'lwc';
import { notifyRecordUpdateAvailable } from 'lightning/uiRecordApi';

export default class RefreshDataOnDisconnect extends LightningElement {

    @api
    recordId

    async disconnectedCallback() {
        if(this.recordId){
            try{
                await notifyRecordUpdateAvailable([{recordId: this.recordId}])
            }
            catch(err){
                window.location.reload();
            }
            
        }
        else{
            window.location.reload();
        }
    }
}