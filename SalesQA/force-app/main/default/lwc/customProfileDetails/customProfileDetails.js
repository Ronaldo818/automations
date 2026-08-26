import { LightningElement } from 'lwc';
import getProfileData from '@salesforce/apex/ProfileDetailsController.getProfileData';

export default class CustomProfileDetails extends LightningElement {
    profileData;
    error;
    errorMessage;

    connectedCallback() {
        this.fetchData();
    }

    fetchData() {
        getProfileData()
            .then((result) => {
                this.profileData = result;
                this.error = undefined;
            })
            .catch((err) => {
                this.error = err;
                this.profileData = undefined;
                this.errorMessage = err.body ? err.body.message : err.message;
            });
    }
}