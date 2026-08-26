import { LightningElement, api, wire } from 'lwc';
import { addItemToCart } from 'commerce/cartApi';
import { NavigationMixin } from 'lightning/navigation';
import basePath from '@salesforce/community/basePath';
import getProducts from '@salesforce/apex/CommerceProductCarousel.getProducts';
import { getSessionContext } from 'commerce/contextApi';
import isGuest from '@salesforce/user/isGuest';

const MOBILE_SLIDE_WIDTH = 150+20;
const DESKTOP_SLIDE_WIDTH = 230+20;
const MOBILE_WIDTH = 600;

export default class CommerceProductCarousel extends NavigationMixin(LightningElement) {
    @api productsSku = null;
    @api exibirPreco = null;
    @api estiloPreco = null;
    @api currencyIsoCode;
    @api exibirAdicionarAoCarrinho = null;
    @api pageSize = 3;
    @api category = null;
    @api categoryProductQuantity = null;
    @api urlPrefix = null;
    @api effectiveAccountId;

    itemsToShow = [];
    currentPage = 1;
    cartId = null;
    showAddedToCart = false;
    slideIndex = 0;
    listaProdutos = [];
    listaProdutosSize = 0;
    previousDisable = true;
    nextDisable = true;
    inSitePreview = false;

    currentPosition = 0;
    currentMargin = 0;
    slidesPerPage = 0;
    currTransl = [];

    // Swipe Up / Down / Left / Right
    initScroll = false;
    initialX = null;
    initialY = null;

    slideWitdth = DESKTOP_SLIDE_WIDTH;
    translationComplete = true;

    addToCartLabel = 'Comprar';

    get foundProducts() {
        return this.listaProdutosSize > 0;
    }

    @wire(getProducts, {categoryName: '$category', productQuantity: '$categoryProductQuantity', stringListSKU: '$productsSku', showPrice: '$exibirPreco', isSiteInPreview: '$inSitePreview', effectiveAccountId: '$effectiveAccountId'})
    wiredGetProducts({ error, data }) {
        if (data) {
            let parsedData = JSON.parse(data);
            this.listaProdutos = parsedData.map((element) => {
                
                // Mapeia possíveis chaves onde o peso possa estar vindo do Apex
                let peso = element.shippingWeight || element.weight || element.ShippingWeight || element.Peso;
                let precoKg = null;

                // Faz o cálculo se os dados estiverem disponíveis
                if (element.price && peso && parseFloat(peso) > 0) {
                    precoKg = element.price / parseFloat(peso);
                }

                return {
                    ...element,
                    defaultImageUrl: basePath + element.defaultImageUrl,
                    precoPorKg: precoKg // Injeta a nova variável no objeto renderizado
                };
            });
            this.listaProdutosSize = this.listaProdutos.length;
            this.checkWidth();
            window.addEventListener("resize", (e) => this.checkWidth(e), false);
        } else if (error) {
            console.error('CommerceProductCarousel wiredGetProducts getProducts:', error);
        }
    }

    async connectedCallback() {
        this.inSitePreview = this.isInSitePreview();
        if(!isGuest){
            await getSessionContext()
                .then(result => {
                    this.effectiveAccountId = result.effectiveAccountId;
                }).catch(error => {
                    console.error('B2bInfiniteCarouselOfLogos getSessionContext: ',error);
            });
        }
    }

    async renderedCallback(){
        if (!this.initScroll) {
            var container = this.template.querySelector('.container');
            container.addEventListener('touchstart', (e) => this.startTouch(e), false);
            container.addEventListener('touchmove', (e) => this.moveTouch(e), false);
            this.initScroll = true;
        }
    }

    disconnectedCallback() {
        if (!this.inSitePreview) {
            try {
                window.removeEventListener("resize", checkWidth);
            } catch (error) {}
        }
    }

    handleAddToCart(event) {
        addItemToCart(event.target.name)
            .then(result => {
                this.cartId = result.cartId;
                this.showAddedToCart = true;
            }).catch(error => {
                console.error(error);
        });
    }

    async checkWidth() {
        await Promise.resolve();
        if (this.refs?.sliderContainer) {
            let w = this.refs.container.offsetWidth;
            this.slideWitdth = w > MOBILE_WIDTH ? DESKTOP_SLIDE_WIDTH : MOBILE_SLIDE_WIDTH;
            if (w < 2*this.slideWitdth) {
                this.slidesPerPage = 1;
            } else if (this.listaProdutosSize-1 < this.pageSize && w > (this.listaProdutosSize-1)*this.slideWitdth) {
                this.slidesPerPage = this.listaProdutosSize-1;
            } else if (this.listaProdutosSize-2 < this.pageSize && w > (this.listaProdutosSize-2)*this.slideWitdth) {
                this.slidesPerPage = this.listaProdutosSize-2;
            } else if (w > this.pageSize*this.slideWitdth) {
                this.slidesPerPage = this.pageSize;
            } else {
                this.slidesPerPage = Math.floor(w/this.slideWitdth);
            }

            if (this.refs.sliderContainer) {
                this.refs.sliderContainer.style.width = (this.slidesPerPage * this.slideWitdth).toString()+'px';
            }

            for(var i = 0; i < this.listaProdutos.length; i++) {
                this.currTransl[i] = -this.slideWitdth;
            }
        }
    }

    get showArrow(){
        return this.slidesPerPage != this.categoryProductQuantity;
    }

    showPrevious() {
        this.translateSlide(-1);
    }

    showNext() {
        this.translateSlide(1);
    }

    translateSlide(factor) {
        if (this.translationComplete === true && (factor == -1 || factor == 1)) {
            this.translationComplete = false;
            this.slideIndex += factor;
            if (this.slideIndex == -1) {
                this.slideIndex = this.listaProdutos.length-1;
            }
            var outerIndex = factor > 0 ? (this.slideIndex-1) % this.listaProdutos.length : (this.slideIndex) % this.listaProdutos.length;

            let slides = this.template.querySelectorAll('.slide');
            for(var i = 0; i < this.listaProdutos.length; i++) {
                var slide = slides[i];
                this.currTransl[i] = this.currTransl[i]-this.slideWitdth*factor;
                slide.style.opacity = '1';
                slide.style.transform = 'translateX('+(this.currTransl[i])+'px)';
            }

            var outerSlide = slides[outerIndex];
            this.currTransl[outerIndex] = this.currTransl[outerIndex]+this.slideWitdth*factor*(this.listaProdutos.length);
            outerSlide.style.opacity = '0';
            outerSlide.style.transform = 'translateX('+(this.currTransl[outerIndex])+'px)';

            setTimeout(() => {
                this.translationComplete = true;
            }, 500);
        }
    }

    closeAddedToCartModal() {
        this.showAddedToCart = false;
    }

    isInSitePreview() {
        let url = document.URL;
        return (
            (url.indexOf("sitepreview") > 0 ||
            url.indexOf("livepreview") > 0 ||
            url.indexOf("live-preview") > 0 ||
            url.indexOf("live.") > 0 ||
            url.indexOf(".builder.") > 0) && !isGuest
        );
    }

    startTouch(e) {
        this.initialX = e.touches[0].clientX;
        this.initialY = e.touches[0].clientY;
    }

    moveTouch(e) {
        if (this.initialX === null) {
            return;
        }

        if (this.initialY === null) {
            return;
        }

        var currentX = e.touches[0].clientX;
        var currentY = e.touches[0].clientY;

        var diffX = this.initialX - currentX;
        var diffY = this.initialY - currentY;

        if (Math.abs(diffX) > Math.abs(diffY)) {
            if (diffX > 0) {
                this.showNext();
            } else {
                this.showPrevious();
            }
        } else {
            if (diffY > 0) {
            } else {
            }
        }

        this.initialX = null;
        this.initialY = null;
        e.preventDefault();
    }

    handleOnMouseEnter(event){
        event.toElement.classList.add('mouse-is-over');
    }   

    handleOnMouseLeave(event){
        event.fromElement.classList.remove('mouse-is-over');
    }
}