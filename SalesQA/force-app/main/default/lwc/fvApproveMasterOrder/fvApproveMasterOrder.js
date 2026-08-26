import { LightningElement, api } from 'lwc';
import getinitData from '@salesforce/apex/MasterOrderApprovalController.getinitData';
import doApproval from '@salesforce/apex/MasterOrderApprovalController.doApproval';
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';

export default class FvApproveMasterOrder extends LightningElement {
    @api recordId;
    isDisabled = true;

    connectedCallback() {
        const toastContainer = ToastContainer.instance();
        toastContainer.maxToasts = 5;
        toastContainer.toastPosition = 'top-center';

        getinitData({ recordId: this.recordId })
        .then(result => {
            this.isDisabled = !result.enableApproval;
        })
        .catch(error => {
            console.error(error);
        });
    }

    handleClick() {
        this.isDisabled = true;

        doApproval({ recordId: this.recordId, status: 'Active' })
        .then(() => {
            Toast.show({
                label: 'Contra proposta aprovada!',
                mode: 'dismissible',
                variant: 'success'
            }, this);
            
        })
        .catch(error => {
            console.error(error);
            this.isDisabled = false;

            Toast.show({
                label: 'Erro ao aprovar contra proposta!',
                mode: 'dismissible',
                variant: 'error'
            }, this);

        });
    }

    handleClickReject() {
        this.isDisabled = true;
        doApproval({ recordId: this.recordId, status: 'Rejected' })
            .then(() => {
                Toast.show({
                    label: 'Contra proposta rejeitada!',
                    mode: 'dismissible',
                    variant: 'success'
                }, this);
                
            })
            .catch(error => {
                console.error(error);
                this.isDisabled = false;

                Toast.show({
                    label: 'Erro ao rejeitar contra proposta!',
                    mode: 'dismissible',
                    variant: 'error'
                }, this);

            });
    }
    
}