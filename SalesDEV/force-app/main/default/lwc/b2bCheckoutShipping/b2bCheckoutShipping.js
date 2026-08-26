import { wire } from 'lwc';

import { CheckoutComponentBase, CheckoutAddressAdapter, updateDeliveryMethod, updateShippingAddress } from 'commerce/checkoutApi';


const CheckoutStage = {
    CHECK_VALIDITY_UPDATE: 'CHECK_VALIDITY_UPDATE',
    REPORT_VALIDITY_SAVE: 'REPORT_VALIDITY_SAVE',
    BEFORE_PAYMENT: 'BEFORE_PAYMENT',
    PAYMENT: 'PAYMENT',
    BEFORE_PLACE_ORDER: 'BEFORE_PLACE_ORDER',
    PLACE_ORDER: 'PLACE_ORDER'
};

const RegionMap ={
    ACRE:'AC',
    ALAGOAS:'AL',
    AMAPA:'AP',
    AMAZONAS:'AM',
    BAHIA:'BA',
    CEARA:'CE',
    DISTRITO_FEDERAL:'DF',
    ESPIRITO_SANTO:'ES',
    GOIAS:'GO',
    MARANHAO:'MA',
    MATO_GROSSO:'MT',
    MATO_GROSSO_DO_SUL:'MS',
    MINAS_GERAIS:'MG',
    PARA:'PA',
    PARAÍBA:'PB',
    PARANÁ:'PR',
    PERNAMBUCO:'PE',
    PIAUI:'PI',
    RIO_DE_JANEIRO:'RJ',
    RIO_GRANDE_DO_NORTE:'RN',
    RIO_GRANDE_DO_SUL:'RS',
    RONDÔNIA:'RO',
    RORAIMA:'RR',
    SANTA_CATARINA:'SC',
    SAO_PAULO:'SP',
    SERGIPE:'SE',
    TOCANTINS:'TO',
}

export default class B2bCheckoutShipping extends CheckoutComponentBase {

    shippingAddress = {};

    @wire(CheckoutAddressAdapter, {addressType: 'Shipping', defaultOnly: true, excludeUnsupportedCountries: false})
    checkoutAddress({ error, data }) {
    if (data) {
        this.shippingAddress = data.items[0];
        console.log('CheckoutAddressAdapter data:',  data);
        updateShippingAddress({
                deliveryAddress: {
                    firstName: data.items[0].firstName, 
                    lastName: data.items[0].lastName, 
                    region: RegionMap[data.items[0].region.replaceAll(' ', '_')], 
                    country: data.items[0].country,
                    city: data.items[0].city,
                    street: data.items[0].street, 
                    postalCode: data.items[0].postalCode 
                }

         })
    } else if (error) {
        console.error("B2bCheckoutShipping CheckoutAddressAdapter: ", error);
    }
    }

    async stageAction(checkoutStage) {
        console.log('B2bCheckoutShipping - ', checkoutStage);
        switch (checkoutStage) {
            default:
                return Promise.resolve(true);
        }
    }


}