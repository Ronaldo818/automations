import { LightningElement, wire } from 'lwc';
import { CartSummaryAdapter, refreshCartSummary } from 'commerce/cartApi';
import expirateCart from '@salesforce/apex/CommerceUtil.expirateCart';

export default class CommerceCustomCartSummary extends LightningElement {

    cartId;
    verbaDisponivel = 0;
    verbaConsumida = 0;
    pesoTotal = 0;
    valorTotal = 0;
    percentualDescontoEscalonado = 0;
    pedidoMesmoDia = false;
    limiteCredito = 0;
    farol = '';

    tempoRestante;
    loadTime;
    loaded = false;
    timeLeft = '00:00:00';

    get showTimer(){
        return (this.timeLeft != null);
    }

    get showOrderAlert(){
        return this.pedidoMesmoDia;
    }

    connectedCallback(){
        this.loadTime = new Date().getTime();
    }

    stageAction(checkoutStage) {
        console.log('CommerceCustomCartSummary-checkoutStage', checkoutStage);

        // switch (checkoutStage) {
        //     case CheckoutStage.BEFORE_PLACE_ORDER:
        //         return Promise.resolve(this.reportValidity());
        //     default:
        //         return Promise.resolve(true);
        // }
    }

    @wire(CartSummaryAdapter, {})
    cartSummary({ error, data }) {
        console.log('cartSummary data', data);
        if (data) {
            this.cartId = data.cartId;
            this.verbaDisponivel = data?.customFields[0]?.SaldoVerbaflex__c ?? 0;
            this.verbaConsumida = data?.customFields[0]?.ValorTotalVerba__c ?? 0;
            this.pesoTotal = data?.customFields[0]?.PesoTotal__c ?? 0;
            this.percentualDescontoEscalonado = (data?.customFields[0]?.PercentualDescontoEscalonado__c ?? 0) / 100;
            this.tempoRestante = (data?.customFields[0]?.TempoRestante__c);
            this.pedidoMesmoDia = (data?.customFields[0]?.PedidoMesmoDiaMesmoCliente__c);
            this.limiteCredito = (data?.customFields[0]?.LimiteCreditoCliente__c);
            this.valorTotal = data.totalProductAmountAfterAdjustments ?? 0;
            this.farol = (data?.customFields[0]?.Farol__c ?? '');

            // if(this.verbaConsumida > 0){
            //     this.farol = 'success';
            // } else if(this.verbaConsumida < 0){
            //     this.farol = 'error';
            // }

            if(!this.loaded){
                this.loaded = true;
                if(this.cartId){
                    this.start();
                }
            }
        } else if (error) {
            console.error(error);
        }
    }

    start(){

        this.tempo = setTimeout(() => {
            var now = new Date().getTime();
            let diff = now - this.loadTime;
            var distance = this.tempoRestante - diff;
            var hours = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
            var minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
            var seconds = Math.floor((distance % (1000 * 60)) / 1000);

            if(hours < 10) hours = '0' + hours;
            if(minutes < 10) minutes = '0' + minutes;
            if(seconds < 10) seconds = '0' + seconds;

            var timeLeft =  hours + ":" + minutes + ":" + seconds;
            this.timeLeft = timeLeft;

            if(distance > 0){
                this.start();
            } else {
                this.timeLeft = '00:00:00';
                this.tempo = setTimeout(() => {
                    expirateCart({
                        cartId: this.cartId
                    }).then(result => {
                        refreshCartSummary().then(() => {
                        }).catch(error => {
                            console.error('refreshCartSummary', error);
                        });
                    }).catch(error => {
                        console.error('expirecart: ', error);
                    })
                }, 100);
            }
        }, 1000);
    }
}