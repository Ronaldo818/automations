import { LightningElement, api } from 'lwc';
import { FlowNavigationNextEvent, FlowNavigationFinishEvent } from 'lightning/flowSupport';

//based on https://github.com/UnofficialSF/LightningFlowComponents/blob/master/flow_screen_components/flowAutoNavigate/force-app/main/default/lwc/flowAutoNavigate/flowAutoNavigate.js

export default class FlowAutoAdvance extends LightningElement {

    @api availableActions = [];

    renderedCallback() {
        var parentThis = this;
        
        // Navigate to the next step in the flow either next action or finish
        if (parentThis.availableActions.find(action => action === 'NEXT')) {
            const navigateNextEvent = new FlowNavigationNextEvent();
            parentThis.dispatchEvent(navigateNextEvent);
        } else if (parentThis.availableActions.find(action => action === 'FINISH')) {
            const navigateFinishEvent = new FlowNavigationFinishEvent();
            parentThis.dispatchEvent(navigateFinishEvent);
        }
    }

}