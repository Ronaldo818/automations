import { LightningElement, api } from 'lwc';
import { ShowToastEvent } from 'lightning/platformShowToastEvent';
import { RefreshEvent } from 'lightning/refresh';

import processPayment from '@salesforce/apex/pixPaymentActionController.processPayment';

export default class PixPaymentActions extends LightningElement {
    @api recordId;
    isProcessing = false;

    async handleCancelar() {
        this.isProcessing = true;
        try {
            await processPayment({ orderId: this.recordId, cieloPaymentStatus: 10});

            this.showToast('Sucesso', 'Pagamento cancelado com sucesso.', 'success');
            this.reloadPage();

        } catch (error) {
            this.handleError(error, 'Erro ao cancelar pagamento');
        } finally {
            this.isProcessing = false;
        }
    }

    async handleReembolsar() {
        this.isProcessing = true;
        try {
            await processPayment({ orderId: this.recordId, cieloPaymentStatus: 11});

            this.showToast('Sucesso', 'Pagamento reembolsado com sucesso.', 'success');
            this.reloadPage();

        } catch (error) {
            this.handleError(error, 'Erro ao reembolsar pagamento');
        } finally {
            this.isProcessing = false;
        }
    }

    async handleConfirmar() {
        this.isProcessing = true;
        try {
            await processPayment({ orderId: this.recordId, cieloPaymentStatus: 2});

            this.showToast('Sucesso', 'Pagamento confirmado com sucesso.', 'success');
            this.reloadPage();

        } catch (error) {
            this.handleError(error, 'Erro ao confirmar pagamento');
        } finally {
            this.isProcessing = false;
        }
    }

    showToast(title, message, variant) {
        this.dispatchEvent(
            new ShowToastEvent({
                title: title,
                message: message,
                variant: variant
            })
        );
    }

    handleError(error, defaultMessage) {
        let message = defaultMessage;

        if (error?.body?.message) {
            message = error.body.message;
        }

        this.showToast('Erro', message, 'error');
        console.error(error);
    }

    reloadPage() {
        // pequeno delay para garantir que o toast apareça antes do reload
        setTimeout(() => {
            this.dispatchEvent(new RefreshEvent());
        }, 1500);
    }
}