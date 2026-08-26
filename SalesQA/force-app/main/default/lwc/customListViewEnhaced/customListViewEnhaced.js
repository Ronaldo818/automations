import { LightningElement, wire } from 'lwc';
import { NavigationMixin } from 'lightning/navigation';
import getAccounts from '@salesforce/apex/CustomListViewEnhacedController.getAccounts';

export default class AccountListPortal extends NavigationMixin(LightningElement) {

    accounts = [];
    filteredAccounts = [];
    searchTerm = '';
    isLoading = true;
    sortBy = 'name';
    sortDirection = 'asc';

    @wire(getAccounts)
    wiredAccounts({ error, data }) {
        this.isLoading = false;

        if(data){
            this.accounts = data.map(acc => ({
                id: acc.id,
                name: acc.name,
                codigoClienteERP: acc.codigoClienteERP,
                cpfCnpj: acc.cpfCnpj,
                categoria: acc.categoriaLabel,
                subcategoria: acc.subcategoriaLabel,
                status: acc.status,
                rowClass: (acc.bloqueio ? 'table-row blocked' : 'table-row'),
                _searchText: [acc.name, acc.codigoClienteERP, acc.cpfCnpj, acc.cleanCpfCnpj, acc.status, acc.categoriaLabel, acc.subcategoriaLabel].filter(Boolean).join(' ').toLowerCase()}));

            this.filteredAccounts = [...this.accounts];
            this.sortData(this.sortBy, this.sortDirection);
        } else if (error) {
            console.error(error);
        }
    }

    handleSearch(event) {
        this.searchTerm = event.target.value.toLowerCase();
        if (!this.searchTerm) {
            this.filteredAccounts = [...this.accounts];
        } else {
            this.filteredAccounts = this.accounts.filter(acc => acc._searchText.includes(this.searchTerm));
        }
        this.sortData(this.sortBy, this.sortDirection);
    }

    handleSort(event) {
        const field = event.currentTarget.dataset.field;
        if(this.sortBy === field){
            this.sortDirection = this.sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
            this.sortBy = field;
            this.sortDirection = 'asc';
        }
        this.sortData(this.sortBy, this.sortDirection);
    }

    sortData(field, direction) {
        const isAsc = direction === 'asc' ? 1 : -1;
        this.filteredAccounts = [...this.filteredAccounts].sort((a, b) => {
            let valueA = a[field] ?? '';
            let valueB = b[field] ?? '';

            if (typeof valueA === 'boolean') {
                return ((valueA === valueB ? 0 : valueA ? 1 : -1) * isAsc);
            }

            valueA = String(valueA).toLowerCase();
            valueB = String(valueB).toLowerCase();

            if (valueA > valueB) return 1 * isAsc;
            if (valueA < valueB) return -1 * isAsc;
            return 0;
        });
    }

    navigateToRecord(event) {
        const recordId = event.currentTarget.dataset.id;
        this[NavigationMixin.Navigate]({
            type: 'standard__recordPage',
            attributes: {
                recordId,
                objectApiName: 'Account',
                actionName: 'view'
            }
        });
    }

    get hasAccounts() {
        return this.filteredAccounts.length > 0;
    }

    getSortIcon(field) {
        if (this.sortBy !== field) {
            return 'utility:';
        } else {
            return this.sortDirection === 'asc' ? 'utility:arrowup' : 'utility:arrowdown';
        }
    }

    get nameSortIcon() {
        return this.getSortIcon('name');
    }

    get codigoSortIcon() {
        return this.getSortIcon('codigoClienteERP');
    }

    get cpfSortIcon() {
        return this.getSortIcon('cpfCnpj');
    }

    get categoriaSortIcon() {
        return this.getSortIcon('categoria');
    }

    get subcategoriaSortIcon() {
        return this.getSortIcon('subcategoria');
    }

    get statusSortIcon() {
        return this.getSortIcon('status');
    }

}