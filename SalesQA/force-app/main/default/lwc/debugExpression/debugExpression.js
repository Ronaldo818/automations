import { LightningElement, api } from 'lwc';

export default class DebugExpression extends LightningElement {


    _expressionData

    @api
    set expressionData(value){
        console.log('Debug Expression:', value);
    }

    get expressionData(){
        return this._expressionData;
    }

    connectedCallback(){
        console.log('connectedCallback:', this._expressionData);
    }

    renderedCallback(){
        console.log('renderedCallback:', this._expressionData);
    }

}