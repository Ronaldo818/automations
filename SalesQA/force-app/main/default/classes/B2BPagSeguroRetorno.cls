public with sharing class B2BPagSeguroRetorno {
    public String id;
    public String reference_id;
    public String created_at;
    public Customer customer;
    public List<Items> items;
    public Shipping shipping;
    public List<Charges> charges;
    public List<String> notification_urls;
    public List<Links> links;
    public Amount amount;
    public List<QrCodes> qr_codes;
    public String logId;
    
    public static B2BPagSeguroRetorno parse(String json) {
        json = json.replaceAll('number', 'number_x');
        json = json.replaceAll('currency', 'currency_x');
        return (B2BPagSeguroRetorno)System.JSON.deserialize(json, B2BPagSeguroRetorno.class);
    }

    public class Customer {
        public String name;
        public String email;
        public String tax_id;
        public List<Phones> phones;
    }

    public class Phones {
        public String type;
        public String country;
        public String area;
        public String number_x;
    }

    public class Items {
        public String reference_id;
        public String name;
        public Integer quantity;
        public Decimal unit_amount;
    }

    public class Shipping {
        public Address address;
    }

    public class Address {
        public String street;
        public String number_x;
        public String complement;
        public String locality;
        public String city;
        public String region_code;
        public String country;
        public String postal_code;
        public String region;
    }

    public class Charges {
        public String id;
        public String reference_id;
        public String status;
        public String created_at;
        public String paid_at;
        public String description;
        public Amount amount;
        public PaymentResponse payment_response;
        public PaymentMethod payment_method;
        public List<Links> links;
    }

    public class Amount {
        public Decimal value;
        public String currency_x;
        public Summary summary;
        public Decimal additional;
        public Decimal discount;
    }

    public class Summary {
        public Decimal total;
        public Decimal paid;
        public Decimal refunded;
    }

    public class PaymentResponse {
        public String code;
        public String message;
        public String reference;
        public RawData raw_data;
    }

    public class RawData {
        public String authorization_code;
        public String nsu;
        public String reason_code;
    }

    public class PaymentMethod {
        public String type;
        public Integer installments;
        public Boolean capture;
        public Card card;
        public Boleto boleto;
        public AuthenticationMethod authentication_method;
        public String soft_descriptor;
    }

    public class Card {
        public String id;    
        public String brand;    
        public String first_digits;     
        public String last_digits;    
        public String exp_month;    
        public String exp_year;
        public Holder holder;
        public Boolean store;
    }

    public class Holder {
        public String name;
        public String tax_id;
        public String email;
        public Address address;
    }

    public class Links {
        public String rel;
        public String href;
        public String media;
        public String type;
    }

    public class Boleto {
        public String id;
        public String barcode;
        public String formatted_barcode;
        public String due_date;
        public InstructionLines instruction_lines;
        public Holder holder;
    }

    public class AuthenticationMethod {
        public String status;
        public String type;
    }

    public class InstructionLines {
        public String line_1;
        public String line_2;
    }

    public class QrCodes {
        public String id;
        public Datetime expiration_date;
        public Amount amount;
        public String text;
        public List<Links> links;
    }
}