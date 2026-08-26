import { LightningElement, wire } from 'lwc';

import { CartAdapter, deleteCurrentCart } from 'commerce/cartApi';
import { SessionContextAdapter } from 'commerce/contextApi';

import getTargetDateTime from '@salesforce/apex/B2BCartCronometerController.getTargetDateTime';

import {NavigationMixin} from "lightning/navigation";


const LOGOUTPAGEREF = {
    type: "comm__loginPage",
    attributes: {
        actionName: "logout"
    }
};


export default class B2BCartCronometer extends NavigationMixin(LightningElement) {

    userId;

    cartId;

    closeCartExecution;

    timeLeft = 0.0;

    hours = '--';

    minutes = '--';

    seconds = '--';

    percentageToFinish = 0;

    totalWaitTime;

    totalHours = 8

    isLoggedIn = false;



    connectedCallback() {
        this.getColecao(this.contentKeys, this.channelName, this.storeName);
    }

    disconnectedCallback() {
        if(this.trocaAutomatica) {
            window.clearInterval(this.closeCartExecution);
        }
    }


    @wire(CartAdapter, {})
    cartInfo({ error, data }) {
        if (data) {
            console.log('CartAdapter data:',  data);
        } else if (error) {
            console.error("B2BCartCronometer CartAdapter: ", error);
        }
    }


    @wire(SessionContextAdapter, {})
    sessionInfo({ error, data }) {
        if (data) {
            console.log('SessionContextAdapter data:',  data);
            getTargetDateTime({userId: data.userId})
            .then((result) => {
                const r = JSON.parse(result);
                console.log('from: ', new Date().toUTCString())
                console.log('to: ', r.targetTime);
                const target = Date.parse(r.targetTime);
                console.log('target:', target);
                console.log('now:', Date.now());
                const milisecondsToExecute = target - Date.now();
                this.closeCartExecution = window.setInterval(() => {
                    deleteCurrentCart()
                    .then((result) => {
                        this[NavigationMixin.Navigate](LOGOUTPAGEREF);
                    }).catch((err) => {
                        this[NavigationMixin.Navigate](LOGOUTPAGEREF);
                    });
                }, milisecondsToExecute)
                this.timeLeft = milisecondsToExecute;
                this.totalWaitTime = r.hours * 1000.0*60.0*60.0;
                this.totalHours = r.hours;
                this.updateCounter();

            }).catch((err) => {
                console.error(err);
            });
            this.isLoggedIn = data.isLoggedIn;
            
        } else if (error) {
            console.error("B2BCartCronometer SessionContextAdapter: ", error);
        }
    }


    updateCounter(){
        this.seconds = (Math.floor((this.timeLeft / 1000.0) % 60.0)).toLocaleString('en-US', {minimumIntegerDigits: 2, maximumFractionDigits: 0, useGrouping: false}); 
        this.minutes = (Math.floor(((this.timeLeft / (1000.0*60.0)) % 60.0))).toLocaleString('en-US', {minimumIntegerDigits: 2, maximumFractionDigits: 0, useGrouping: false});
        this.hours   = (Math.floor(((this.timeLeft / (1000.0*60.0*60.0)) % 24.0))).toLocaleString('en-US', {minimumIntegerDigits: 2, maximumFractionDigits: 0, useGrouping: false});

        this.timeLeft -= 1000.0;

        //this.percentageToFinish = 100 - (100 * this.timeLeft)/this.totalWaitTime;
        
        setTimeout(() => {
            if(timeLeft > 0){
                this.updateCounter();
            }
            else{
                this.seconds = '00';
                this.minutes = '00'; 
                this.hours = '00';
            }
        }, 1000);
    }
}