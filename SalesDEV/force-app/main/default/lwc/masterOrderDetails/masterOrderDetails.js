import { LightningElement, wire, api } from 'lwc';
import getMasterOrder from '@salesforce/apex/MasterOrderDetailDataProvider.getMasterOrder';
import { effectiveAccount } from 'commerce/effectiveAccountApi'; 
export default class MasterOrderDetails extends LightningElement {
    @api
    recordId

    effectiveAccount = effectiveAccount;

    masterOrder = undefined

    fetchingData = false
    /*
    @wire(getMasterOrder, { recordId: "$recordId", effectiveAccountId: "$effectiveAccount.accountId" })
    wiredMasterOrder({ data, error }) {
        if(data){
            this.masterOrder = data;
            console.log(this.masterOrder);
        }
        else if(error){
            console.error(error);
        }
    }
    */
    renderedCallback(){
        if(!this.masterOrder && this.recordId && this.effectiveAccount.accountId && !this.fetchingData){
            this.fetchingData = true;
            getMasterOrder({ recordId: this.recordId, effectiveAccountId: this.effectiveAccount.accountId})
            .then(result => {
                this.masterOrder = result;
                this.fetchingData = false;
            })
            .catch(error => {
                console.error('getMasterOrder:', error);
                this.fetchingData = false;
            })
        }
    }
}