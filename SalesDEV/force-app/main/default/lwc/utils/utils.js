import { LightningElement } from 'lwc';
import Toast from 'lightning/toast';
import ToastContainer from 'lightning/toastContainer';

export default class Utils extends LightningElement {
        connectedCallback() {

    
            const toastContainer = ToastContainer.instance();
            toastContainer.maxToasts = 5;
            toastContainer.toastPosition = 'top-center';
        }
}

export function validateEmail(email){
    var regex = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return regex.test(email);
}

export function maskPhone(phone){
    var value = phone.replace(/\D+/g, '');
    var x;
    if(value.length <= 10){
        x = value.match(/(\d{0,2})(\d{0,4})(\d{0,4})/);
    } else {
        x = value.match(/(\d{0,2})(\d{0,5})(\d{0,4})/);
    }
    return !x[2] ? x[1] : `(${x[1]}) ${x[2]}` + (x[3] ? `-${x[3]}` : ``);
}

export function maskCep(cep){
    var cleanValue = cep.replace(/\D+/g, '');
    var x = cleanValue.match(/(\d{0,5})(\d{0,3})/);
    return !x[2] ? x[1] : `${x[1]}-${x[2]}`;
}

export function validateCep(cep){
    var regex = /^\d{8}$/;
    return regex.test(cep.replace(/\D+/g, ''));
}

export function maskCpf(cpf){
    var cleanValue = cpf.replace(/\D+/g, '');
    var x = cleanValue.match(/(\d{0,3})(\d{0,3})(\d{0,3})(\d{0,2})/);

    var maskedValue = '';
    if(x[1]) maskedValue += `${x[1]}`;
    if(x[2]) maskedValue += `.${x[2]}`;
    if(x[3]) maskedValue += `.${x[3]}`;
    if(x[4]) maskedValue += `-${x[4]}`;
    return maskedValue;
}

export function validateCpf(cpf){
    let strCPF = cpf.replace(/[^\d]+/g,'');	
    if(strCPF == '') return false;	
    
    var Soma, Resto;
    Soma = 0;
    if (strCPF == "00000000000" ||
        strCPF == "11111111111" ||
        strCPF == "22222222222" ||
        strCPF == "33333333333" ||
        strCPF == "44444444444" ||
        strCPF == "55555555555" ||
        strCPF == "66666666666" ||
        strCPF == "77777777777" ||
        strCPF == "88888888888" ||
        strCPF == "99999999999") return false;
    
    for (var i=1; i<=9; i++) Soma = Soma + parseInt(strCPF.substring(i-1, i)) * (11 - i);
    Resto = (Soma * 10) % 11;
    
    if ((Resto == 10) || (Resto == 11))  Resto = 0;
    if (Resto != parseInt(strCPF.substring(9, 10)) ) {
        return false;
    }
    
    Soma = 0;
    for (var i = 1; i <= 10; i++) Soma = Soma + parseInt(strCPF.substring(i-1, i)) * (12 - i);
    Resto = (Soma * 10) % 11;
    if ((Resto == 10) || (Resto == 11))  Resto = 0;
    if (Resto != parseInt(strCPF.substring(10, 11) ) ) return false;
    return true;
}

export function validateBirthDate(date){
    if(date){ 
        let parts = date.split('-');
        let dataNascimento = new Date(parts[0], parts[1] - 1, parts[2]);
        let gmtToday = new Date();
        gmtToday.setHours(0, 0, 0, 0);

        if (dataNascimento >= gmtToday) {
            return false;
        } else {
            return true;
        }
    } else {
        return false;
    }
}

export function maskCnpj(cnpj){
    const cleanValue = cnpj.replace(/[^A-Z0-9a-z]/g, '').toUpperCase()
    const x = cleanValue.match(/([A-Z0-9]{0,2})([A-Z0-9]{0,3})([A-Z0-9]{0,3})([A-Z0-9]{0,4})([A-Z0-9]{0,2})/)
    let maskedValue = ''
    if(x[1]) maskedValue += `${x[1]}`
    if(x[2]) maskedValue += `.${x[2]}`
    if(x[3]) maskedValue += `.${x[3]}`
    if(x[4]) maskedValue += `/${x[4]}`
    if(x[5]) maskedValue += `-${x[5]}`
    return maskedValue
}


export function validateCnpj(cnpj){
    const clean = cnpj.replace(/[^A-Z0-9a-z]/g, '').toUpperCase()
    if(clean.length !== 14) return false
    if(new Set(clean.split('')).size === 1) return false

    const mult1 = [5,4,3,2,9,8,7,6,5,4,3,2]
    const mult2 = [6,5,4,3,2,9,8,7,6,5,4,3,2]
    let sum1 = 0, sum2 = 0

    for(let i = 0; i < 12; i++){
        const val = clean.charCodeAt(i) - 48
        sum1 += val * mult1[i]
        sum2 += val * mult2[i]
    }

    const resto1 = sum1 % 11
    const dv1 = resto1 < 2 ? 0 : 11 - resto1

    sum2 += dv1 * mult2[12]
    const resto2 = sum2 % 11
    const dv2 = resto2 < 2 ? 0 : 11 - resto2

    return clean[12] === String(dv1) && clean[13] === String(dv2)
}

export function isInSitePreview() {
    let url = document.URL;
    return (
        url.indexOf("sitepreview") > 0 ||
        url.indexOf("livepreview") > 0 ||
        url.indexOf("live-preview") > 0 ||
        url.indexOf("live.") > 0 ||
        url.indexOf(".builder.") > 0
    );
}