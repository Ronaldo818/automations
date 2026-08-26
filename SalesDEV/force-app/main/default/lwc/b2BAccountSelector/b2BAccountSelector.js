import { LightningElement, track } from 'lwc';
import { effectiveAccount } from 'commerce/effectiveAccountApi';
import getRelatedAccounts from '@salesforce/apex/B2BAccountSelector.getRelatedAccounts';
import getCurrentUserAccount from '@salesforce/apex/B2BAccountSelector.getCurrentUserAccount';
import USER_ID from '@salesforce/user/Id';
import { NavigationMixin } from 'lightning/navigation';

export default class B2BAccountSelector extends NavigationMixin(LightningElement) {
    isModalOpen = false;
    relatedAccounts = [];
    filteredAccounts = [];
    currentAccount = null;
    selectedAccount = null;
    accountName = '';
    isLoading = false;
    loaded = false;

    get noAccounts() {
        return (this.loaded && (this.relatedAccounts.length === 0 || this.filteredAccounts.length === 0));
    }

    get bloqueio() {
        return (this.accountName && this.accountName.startsWith("(BLOQUEADO) "));
    }

    connectedCallback() {
        if (USER_ID) {
            setTimeout(() => {
                this.loadRelatedAccounts();
            }, 3000); 
        }
        this.accountName = effectiveAccount.accountName;
    }

    async loadRelatedAccounts() {
        try {
            this.isLoading = true;
            this.currentAccount = await getCurrentUserAccount({ userId: USER_ID });
            const effectiveAccountId = effectiveAccount.accountId;
            this.selectedAccount = effectiveAccountId || this.currentAccount?.Id;
            const result = await getRelatedAccounts({ userId: USER_ID });
        
            if (result && result.length > 0){;
                const normalizedSelectedId = (this.selectedAccount || '').trim().toLowerCase();
                const accounts = result.map(account => {
                    const normalizedAccountId = (account.accountId || '').trim().toLowerCase();
                    const isSelected = normalizedAccountId === normalizedSelectedId;
                    return {...account, isSelected
                    };
                });

                this.relatedAccounts = accounts;
                this.filteredAccounts = accounts;
                this.loaded = true;
                // if (!effectiveAccount.accountId) {
                    // this.isModalOpen = true;
                // }
            }
        } catch (e) {
            console.error('getRelatedAccounts: ', e);
        } finally {
            this.isLoading = false;
        }
    }

    handleOpenModal() {
        this.searchTerm = '';
        this.filteredAccounts = this.relatedAccounts;
        this.isModalOpen = true;
    }

    handleClose() {
        this.isModalOpen = false;
    }

    handleSearch(event) {
        let searchTerm = event.target.value;
            this.filteredAccounts = this.relatedAccounts.filter(account => {
                return (
                    (account?.name.toLowerCase().includes(searchTerm.toLowerCase())) ||
                    (account?.codigoClienteERP.includes(searchTerm)) ||
                    (account?.cpfCnpj.includes(searchTerm)) ||
                    (account?.cleanCpfCnpj.includes(searchTerm))
                );
            });
    }

    handleListClick(event){
        const accountId = event.currentTarget.dataset.id;
        const selected = this.relatedAccounts.find(account => account.accountId === accountId);
        if (accountId) {
            effectiveAccount.update(accountId, selected.name);
            this.accountName = selected.name;
            this.isModalOpen = false;

            setTimeout(() => {
                this[NavigationMixin.Navigate]({
                    type: 'comm__namedPage',
                    attributes: {
                        name: 'Home'
                    },
                });
                setTimeout(() => {
                    window.location.reload()
                }, 250);
            }, 250);
        }

    }

}