import { LightningElement, api } from 'lwc';
import { deleteCurrentCart } from 'commerce/cartApi';

export default class CommerceCustomEmptyCartButton extends LightningElement {
    
    @api variant;
    @api label;

    handleEmpty(){
        deleteCurrentCart().then(() => {

        }).catch(error => {
            console.error('deleteCurrentCart', error);
        });
    }

}