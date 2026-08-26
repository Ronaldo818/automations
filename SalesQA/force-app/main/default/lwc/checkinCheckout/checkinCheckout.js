import { LightningElement, api, wire } from 'lwc';
import { refreshApex } from '@salesforce/apex';
import { getRecordNotifyChange } from 'lightning/uiRecordApi';
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';
import getStatus from '@salesforce/apex/CheckinCheckoutController.getStatus';
import doCheckin from '@salesforce/apex/CheckinCheckoutController.doCheckin';
import doCheckout from '@salesforce/apex/CheckinCheckoutController.doCheckout';

export default class CheckinCheckout extends LightningElement {

    @api recordId;

    showSpinner = false;
    disableButton = true;
    lat = 0;
    lon = 0;
    hasCheckin = false;
    hasCheckout = false;
    _wiredStatus;

    // ===== Campos do Checkin (Tipo / Contato / Comentário) ======================
    // Mantidos comentados conforme combinado — prontos para reativar.
    //
    // tipoSelectedValue = 'Negociação';
    // tipoOptions = [
    //     { label: 'Negociação',     value: 'Negociação' },
    //     { label: 'Prospecção',     value: 'Prospecção' },
    //     { label: 'Visita Técnica', value: 'Visita Técnica' }
    // ];
    // contatoSelectedValue = null;
    // contatoOptions = [];
    // comentario = '';
    //
    // get isCheckinPhase()  { return !this.hasCheckin; }
    // get hasContatos()     { return this.contatoOptions && this.contatoOptions.length > 0; }
    // handleTipoChange(event)       { this.tipoSelectedValue    = event.detail.value; }
    // handleContatoChange(event)    { this.contatoSelectedValue = event.detail.value; }
    // handleComentarioChange(event) { this.comentario           = event.detail.value; }
    // ============================================================================

    @wire(getStatus, { recordId: '$recordId' })
    wiredStatus(result) {
        this._wiredStatus = result;
        if (result.data) {
            this.hasCheckin  = result.data.hasCheckin;
            this.hasCheckout = result.data.hasCheckout;
        }
    }

    get buttonLabel() {
        if (this.isCompleted) return 'Visita Concluída';
        return this.hasCheckin ? 'Checkout' : 'Checkin';
    }

    get buttonClass() {
        const base = 'checkin-button slds-button';
        return this.isCompleted
            ? `${base} slds-button_success`
            : `${base} slds-button_brand`;
    }

    get isCompleted() {
        return this.hasCheckin && this.hasCheckout;
    }

    connectedCallback() {
        const toastContainer = ToastContainer.instance();
        toastContainer.maxToasts = 5;
        toastContainer.toastPosition = 'top-center';

        this.loadPosition();
    }

    loadPosition() {
        if (this.isCompleted) {
            this.disableButton = true;
            return;
        }

        if (!navigator.geolocation) {
            this.showToast('Geolocalização indisponível', '', 'error');
            return;
        }

        navigator.geolocation.getCurrentPosition(
            (position) => {
                this.lat = position.coords.latitude;
                this.lon = position.coords.longitude;
                this.disableButton = this.isCompleted;
            },
            (error) => {
                this.disableButton = this.isCompleted;
                this.showToast('Erro na geolocalização', this.geolocationMessage(error), 'error');
            },
            { enableHighAccuracy: true, timeout: 5000, maximumAge: 0 }
        );
    }

    geolocationMessage(error) {
        switch (error.code) {
            case error.PERMISSION_DENIED:    return 'Requisição de geolocalização negada para o usuário.';
            case error.POSITION_UNAVAILABLE: return 'Informação de geolocalização indisponível.';
            case error.TIMEOUT:              return 'O tempo da requisição expirou.';
            default:                         return 'Ocorreu um erro desconhecido.';
        }
    }

    async handleClick() {
        if (this.isCompleted) return;

        this.showSpinner = true;
        try {
            if (!this.hasCheckin) {
                await doCheckin({ recordId: this.recordId, lat: this.lat, lon: this.lon });
                this.showToast('Checkin registrado', '', 'success');
            } else {
                // O checkout só conclui se o "Resultado da Visita" estiver preenchido
                // no registro — a regra é validada no servidor (doCheckout).
                await doCheckout({ recordId: this.recordId, lat: this.lat, lon: this.lon });
                this.showToast('Checkout registrado', '', 'success');
            }
            getRecordNotifyChange([{ recordId: this.recordId }]);
            await refreshApex(this._wiredStatus);
            if (this.isCompleted) this.disableButton = true;
        } catch (error) {
            this.showToastError('Erro ao registrar', error);
        } finally {
            this.showSpinner = false;
        }
    }

    showToast(label, message, variant) {
        Toast.show({ label, message, variant, mode: 'dismissible' }, this);
    }

    showToastError(label, error) {
        console.error(error);
        let message = '';
        if (error && error.message)                   message = error.message;
        else if (error && error.body && error.body.message) message = error.body.message;
        Toast.show({ label, message, variant: 'error', mode: 'dismissible' }, this);
    }
}