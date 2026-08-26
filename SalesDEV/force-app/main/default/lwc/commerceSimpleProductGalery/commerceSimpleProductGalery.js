import { LightningElement, api, track, wire } from 'lwc';
import basePath from '@salesforce/community/basePath';

import FORM_FACTOR from '@salesforce/client/formFactor';

import chevron_l from "@salesforce/resourceUrl/chevron_left";

const EXIBIR_CARD = 'slds-show';
const OCULTAR_CARD = 'slds-hide';
const EXIBIR_CIRCULO = 'circulo-style exibir-circulo-style';
const OCULTAR_CIRCULO = 'circulo-style';
const EXIBIR_THUMB = 'thumb-style exibir-thumb-style image-container';
const OCULTAR_THUMB = 'thumb-style image-container';
const SCROLL_TIME = 5000;

export default class CommerceCustomCarouselEnhanced extends LightningElement {
    _mediaGroups
    @api
    set mediaGroups(value) {
        if (value) {
            console.log('set mediaGroups: ', value);
            this._mediaGroups = value;
            this.setupImages();
        }
    }

    async setupImages() {
        this._mediaGroups.forEach(element => {
            if (element.developerName == 'productDetailImage' && element.mediaItems) {
                this.colecoes = element.mediaItems.map((item, index) => {
                    return index === 0 ? {
                        id: item.id,
                        title: item.title,
                        url: basePath + '/sfsites/c' + item.url,
                        slideIndex: index + 1,
                        cardClass: EXIBIR_CARD,
                        dotClass: EXIBIR_CIRCULO,
                        thumbClass: EXIBIR_THUMB

                    } : {
                        id: item.id,
                        title: item.title,
                        url: basePath + '/sfsites/c' + item.url,
                        slideIndex: index + 1,
                        cardClass: OCULTAR_CARD,
                        dotClass: OCULTAR_CIRCULO,
                        thumbClass: OCULTAR_THUMB
                    }
                });

                this.conteudo = this.colecoes != null ? true : false;
                this.exibirBotoes = this.colecoes.length > 1 ? true : false;
                this.exibirThumbs = this.exibirThumbs && this.colecoes.length > 1;

                this.flushPromises()
                    .then((result) => {
                        // loading das imanges
                        const images = this.template.querySelectorAll('.image-container');
                        images.forEach(div => {
                            const img = div.querySelector('img');

                            function loaded() {
                                div.classList.add('loaded');
                            }

                            if (img.complete) {
                                loaded();
                            } else {
                                img.addEventListener('load', loaded);
                            }
                        })
                        this.isLoading = false;
                    }).catch((err) => {
                        console.error(err);
                        this.isLoading = false;
                    });
            }
        })
    }

    get mediaGroups() {
        return this._mediaGroups
    }

    @api slideTime = SCROLL_TIME;
    @api trocaAutomatica;
    @api exibirSetas;
    @api exibirCirculos;
    @api exibirThumbs;
    @api trocaManual;
    @api zoomEnable;
    @api zoomTop;
    @api recordId;

    exibirBotoes;
    conteudo = true;
    effectiveAccountId = null;
    isLoading = true;
    timer;

    lens;

    carouselSlideIndex = 1;
    slideIndex = 1;
    colecoes = [{
        title: 'loading',
        url: '',
        slideIndex: 1,
        cardClass: EXIBIR_CARD,
        dotClass: EXIBIR_CIRCULO,
        thumbClass: EXIBIR_THUMB
    }];

    showButton = false;

    chevron_l_src = chevron_l + '#Layer_1';

    get isMobile() {
        return FORM_FACTOR != 'Large';
    }


    get buttonStyle() {

        return 'background-color:rgba(255, 255, 255, 0.6);height:100%;width:10%;display:flex;justify-content:center;align-items:center;position:unset'

    }

    renderedCallback() {
        if (this.isMobile) {
            this.showButton = true;
        }
    }

    async flushPromises() {
        return Promise.resolve();
    }

    disconnectedCallback() {
        if (this.trocaAutomatica) {
            window.clearInterval(this.timer);
        }
    }

    prevCard() {
        if (this.trocaManual && this.exibirSetas) {
            const slideIndex = this.slideIndex - 1;
            this.slideSelectionHandler(slideIndex);
        }
    }

    nextCard() {
        if (this.trocaManual && this.exibirSetas) {
            const slideIndex = this.slideIndex + 1;
            this.slideSelectionHandler(slideIndex);
        }
    }

    selectCard(event) {
        if (this.trocaManual) {
            const slideIndex = Number(event.target.dataset.id);
            this.slideSelectionHandler(slideIndex);
        }
    }

    prevCardCarousel() {
        if (this.trocaManual && this.exibirSetas) {
            const slideIndex = this.carouselSlideIndex - 1;
            this.carouselSelectionHandler(slideIndex);
        }
    }

    nextCardCarousel() {
        if (this.trocaManual && this.exibirSetas) {
            const slideIndex = this.carouselSlideIndex + 1;
            this.carouselSelectionHandler(slideIndex);
        }
    }

    carouselSelectionHandler(index) {
        if (index > this.colecoes.length) {
            this.carouselSlideIndex = 1;
        } else if (index < 1) {
            this.carouselSlideIndex = this.colecoes.length;
        } else {
            this.carouselSlideIndex = index;
        }

        if (this.exibirThumbs) {
            this.refs.scrollPanel.style.transform = "translateX(" + ((this.carouselSlideIndex - 1) * -80) + "px)";
        }
    }

    slideSelectionHandler(index) {
        if (index > this.colecoes.length) {
            this.slideIndex = 1;
        } else if (index < 1) {
            this.slideIndex = this.colecoes.length;
        } else {
            this.slideIndex = index;
        }

        // preload image in zoom window for smooth enter
        if (this.zoomEnable) {
            this.refs.zoomArea.style.backgroundImage = "url('" + this.colecoes[this.slideIndex - 1].url + "')";
        }

        this.carouselSelectionHandler(this.slideIndex);

        this.colecoes = this.colecoes.map(item => {
            return this.carouselSlideIndex === item.slideIndex ? {
                ...item,
                cardClass: EXIBIR_CARD,
                dotClass: EXIBIR_CIRCULO,
                thumbClass: EXIBIR_THUMB

            } : {
                ...item,
                cardClass: OCULTAR_CARD,
                dotClass: OCULTAR_CIRCULO,
                thumbClass: OCULTAR_THUMB
            }
        });
    }

    showButtons() {
        this.showButton = true;
    }

    hideButtons() {
        if (!this.isMobile) {
            this.showButton = false
        }
    }

    handleDownloadImage() { this.downloadImage('.png'); }


    downloadImage(extension) {
        if (this.slideIndex && this.colecoes[this.slideIndex - 1] && this.colecoes[this.slideIndex - 1].url) {
            // Faz download da imagem
            const link = document.createElement('a');
            link.href = this.colecoes[this.slideIndex - 1].url;
            link.download = this.colecoes[this.slideIndex - 1]?.title + extension;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }
    }

    _showZoom = false;

    // show and hide zoom window
    set showZoom(value) {
        if (value != this._showZoom) {
            this._showZoom = value;
            if (value) {
                this.refs.zoomArea.className = this.refs.zoomArea.className.replace('slds-hide', 'slds-show');
            }
            else {
                // remove the lens, it is added again when the new zoom is created
                this.lens.remove();
                this.refs.zoomArea.className = this.refs.zoomArea.className.replace('slds-show', 'slds-hide');
            }
        }
    }

    get showZoom() {
        return this._showZoom;
    }



    imageZoom(imgElement, resultClass) {
        var img, result, localLens, cx, cy;

        function getCursorPos(e) {
            var a, x = 0, y = 0;
            e = e;
            /* Get the x and y positions of the image: */
            a = img.getBoundingClientRect();
            /* Calculate the cursor's x and y coordinates, relative to the image: */
            x = e.pageX - a.left;
            y = e.pageY - a.top;
            /* Consider any page scrolling: */
            x = x - window.scrollX;
            y = y - window.scrollY;
            return { x: x, y: y };
        }

        function moveLens(e) {
            //console.log('localLens moving! ', e);

            var pos, x, y;
            /* Prevent any other actions that may occur when moving over the image */
            e.preventDefault();
            /* Get the cursor's x and y positions: */
            pos = getCursorPos(e);
            /* Calculate the position of the localLens: */
            x = pos.x - (localLens.offsetWidth / 2);
            y = pos.y - (localLens.offsetHeight / 2);
            /* Prevent the localLens from being positioned outside the image: */
            if (x > img.width - localLens.offsetWidth) { x = img.width - localLens.offsetWidth; }
            if (x < 0) { x = 0; }
            if (y > img.height - localLens.offsetHeight) { y = img.height - localLens.offsetHeight; }
            if (y < 0) { y = 0; }
            /* Set the position of the localLens: */
            localLens.style.left = x + "px";
            localLens.style.top = y + "px";
            /* Display what the localLens "sees": */
            result.style.backgroundPosition = "-" + (x * cx) + "px -" + (y * cy) + "px";
        }

        img = imgElement;
        this.flushPromises();
        result = this.template.querySelector('.' + resultClass);

        // set zoom window the same size as the image
        result.style.height = img.height + 'px';
        result.style.width = img.width + 'px';

        result.style.top = this.zoomTop + 'px';

        /* Create this.lens: */
        if (this.lens) {
            this.lens.remove();
        }
        this.lens = document.createElement("DIV");
        this.lens.setAttribute('c-commercecustomcarouselenhanced_commercecustomcarouselenhanced', true);
        this.lens.setAttribute("class", "img-zoom-this.lens");
        this.lens.setAttribute("style", "position: absolute; border: 1px solid #d4d4d4;z-index: 8;")
        this.lens.onmouseleave = this.undisplayZoom;

        let computed = getComputedStyle(result, null);
        let w = computed.getPropertyValue('width');
        let h = computed.getPropertyValue('height');

        //console.log('result height:', h);
        //console.log('result width:', w);

        //console.log(parseFloat(h.replace('px','')) /5.0 + 'px');
        //console.log(parseFloat(w.replace('px','')) /5.0 + 'px');

        this.lens.style.height = parseFloat(h.replace('px', '')) / 5.0 + 'px';
        this.lens.style.width = parseFloat(w.replace('px', '')) / 5.0 + 'px';
        /* Insert this.lens: */
        img.parentElement.insertBefore(this.lens, img);

        /* Calculate the ratio between result DIV and this.lens: */
        cx = result.offsetWidth / this.lens.offsetWidth;
        cy = result.offsetHeight / Math.max(this.lens.offsetHeight, 1);
        /* Set background properties for the result DIV */
        result.style.backgroundImage = "url('" + img.src + "')";
        result.style.backgroundSize = (img.width * cx) + "px " + (img.height * cy) + "px";
        /* Execute a function when someone moves the cursor over the image, or the this.lens: */
        localLens = this.lens;
        this.lens.addEventListener("mousemove", moveLens);
        img.addEventListener("mousemove", moveLens);
        /* And also for touch screens: */
        this.lens.addEventListener("touchmove", moveLens);
        img.addEventListener("touchmove", moveLens);
    }

    displayZoom(event) {
        // display zoom only if the zoom is enable and image is loades
        if (this.zoomEnable && !this.showZoom && event.toElement.querySelector('.loaded')) {
            this.showZoom = true;
            this.imageZoom(event.toElement.querySelector('img'), 'zoomArea')
        }
    }

    undisplayZoom(event) {
        // undisplay zoom only if is not going to the lens element
        if (this.zoomEnable && this.showZoom) {
            if (!event.toElement || !event.toElement.classList.contains('img-zoom-this.lens')) {
                this.showZoom = false;
            }

        }
    }
}