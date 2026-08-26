import { LightningElement, api, wire } from 'lwc';
import getLocations from '@salesforce/apex/CheckinCheckoutMapController.getLocations';

export default class CheckinCheckoutMap extends LightningElement {

    @api recordId;

    // Configuráveis via App Builder (ver meta.xml)
    @api cardTitle = 'Mapa da Visita';
    @api zoomLevel = 15;
    @api listView  = 'visible'; // 'visible' | 'hidden' | 'auto'

    mapMarkers   = [];
    center;
    markersTitle = 'Localizações registradas';
    error;
    isLoading    = true;

    @wire(getLocations, { recordId: '$recordId' })
    wiredLocations({ data, error }) {
        this.isLoading = false;
        if (data) {
            this.error = undefined;
            this.buildMarkers(data);
        } else if (error) {
            this.error = error;
            this.mapMarkers = [];
        }
    }

    buildMarkers(data) {
        const markers = [];

        if (data.latCheckin != null && data.lonCheckin != null) {
            markers.push({
                location: {
                    Latitude:  data.latCheckin,
                    Longitude: data.lonCheckin
                },
                title: 'Checkin',
                description: this.buildDescription('Início da visita', data.horaCheckin),
                icon: 'utility:location'
            });
        }

        if (data.latCheckout != null && data.lonCheckout != null) {
            markers.push({
                location: {
                    Latitude:  data.latCheckout,
                    Longitude: data.lonCheckout
                },
                title: 'Checkout',
                description: this.buildDescription('Fim da visita', data.horaCheckout),
                icon: 'utility:check'
            });
        }

        this.mapMarkers = markers;

        // Centraliza no primeiro marcador para garantir foco mesmo com 1 só ponto
        if (markers.length > 0) {
            this.center = { location: markers[0].location };
        }
    }

    buildDescription(label, dt) {
        if (!dt) return label;
        const formatted = new Date(dt).toLocaleString('pt-BR', {
            day: '2-digit', month: '2-digit', year: 'numeric',
            hour: '2-digit', minute: '2-digit'
        });
        return `<b>${label}</b><br/>${formatted}`;
    }

    get hasMarkers() {
        return this.mapMarkers && this.mapMarkers.length > 0;
    }

    get hasError() {
        return this.error != null;
    }
}