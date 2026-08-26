public with sharing class B2BPagSeguroRequisicao {
    public String reference_id;
    public Customer customer;
    public List<Items> items;
    public Shipping shipping;
    public List<String> notification_urls;
    public List<Charges> charges;
    public List<QrCodes> qr_codes;

    public Static String parse(B2BPagSeguro.ApexInput input, WebCart cart) {
        Account account = input.account;

        Contact contact = [
            SELECT Id, Name, Email, AccountId, MobilePhone, Phone, CPF__c, DDD_MobilePhone__c, DDD_Phone__c
            FROM Contact
            WHERE Id IN (SELECT ContactId FROM User WHERE Id =: UserInfo.getUserId())
        ];
    
        List<CartItem> cis = [
            SELECT Id, Product2.Name, TotalPrice, Quantity, Product2.Description, Product2.StockKeepingUnit 
			FROM CartItem
            WHERE CartId = :cart.Id
        ];

        WebService__mdt service = WebService__mdt.getInstance('PagSeguro');
    
        B2BPagSeguroRequisicao b2bRequisicao = new B2BPagSeguroRequisicao();
        b2bRequisicao.reference_id = cart.Id;
        b2bRequisicao.customer = new Customer();
        b2bRequisicao.items = new List<Items>();
        b2bRequisicao.shipping = new Shipping();
        b2bRequisicao.notification_urls = new List<String>();
        b2bRequisicao.notification_urls.add(service.NotificationURL__c);
    
        // Customer Information
        Customer cliente = new Customer();
        cliente.name = contact.Name;
        cliente.email = contact.Email;
        cliente.tax_id = contact.CPF__c;
	    cliente.phones = new List<Phones>();
        if(contact.MobilePhone != null) {
        	Phones mobile = new Phones();
        	mobile.country = '55'; // Default Brazil
        	mobile.area = contact.DDD_MobilePhone__c;
        	mobile.number_x = contact.MobilePhone;
        	mobile.type = 'MOBILE';
        	cliente.phones.add(mobile);
        }
        if(contact.Phone != null) {
        	Phones phone = new Phones();
        	phone.country = '55'; // Default Brazil
        	phone.area = contact.DDD_Phone__c;
        	phone.number_x = contact.Phone;
        	phone.type = 'BUSINESS';  
        	cliente.phones.add(phone);
        }
        b2bRequisicao.customer = cliente;
    
        // Product Informations
        for(CartItem produto : cis) {
        	Items item = new Items();
        	item.reference_id = produto.Id;
        	item.name = produto.Product2.Name;
        	item.quantity = Integer.valueOf(produto.quantity);
        	item.unit_amount = (produto.TotalPrice / produto.Quantity * 100).setScale(0);
        	b2bRequisicao.items.add(item);
        }
    
        // Delivery Informations
        Shipping endereco = new Shipping();
        endereco.address = new Address();
        endereco.address.street = account.RuaEntrega__c;
        endereco.address.number_x = account.NumeroEntrega__c;
        endereco.address.complement = account.ComplementoEntrega__c;
        endereco.address.locality = account.BairroEntrega__c;
        endereco.address.postal_code = account.CEPEntrega__c.replaceAll('[^0-9.]', '');
        endereco.address.city = account.CidadeEntrega__r.Name;
        endereco.address.region_code = account.CidadeEntrega__r.UF__c;
        endereco.address.country = 'BRA';
        b2bRequisicao.shipping = endereco;
        
        // Order Value
        Amount quantia = new Amount();
        quantia.value = ((cart.GrandTotalAmount - cart.ValorCashback__c) * 100).setScale(0);
        quantia.currency_x = 'BRL';

        // Payment Informations
        if(input.formaPagamento == 'pix') {
        	b2bRequisicao.qr_codes = new List<QrCodes>();
        	QrCodes qrCodes = new QrCodes();
            StandardValue__mdt pixHours = [SELECT Value__c FROM StandardValue__mdt WHERE DeveloperName = 'PixHoursToExpire' LIMIT 1];
        	Datetime validade = Datetime.now().addHours(Integer.valueOf(pixHours.Value__c));
        	qrCodes.expiration_date = validade.formatGmt('yyyy-MM-dd\'T\'HH:mm:ss') + '+00:00';
        	qrCodes.amount = new Amount();
        	qrCodes.amount = quantia;
        	b2bRequisicao.qr_codes.add(qrCodes);
        }
        else {
        	b2bRequisicao.charges = new List<Charges>();
        	Charges charg = new Charges();
        	charg.reference_id = cart.Id;
        	charg.description = 'Compra Avivar';
        	charg.amount = new Amount();
        	charg.amount = quantia;
        	if(input.formaPagamento.equalsIgnoreCase('cc')) {
            	PaymentMethod pagamento = new PaymentMethod();
            	pagamento.type = 'CREDIT_CARD';
            	pagamento.installments = Integer.valueOf(input.parcelas.replaceAll('[^0-9]', ''));
            	pagamento.capture = true;
            	pagamento.card = new Card();
            	Card cartao = new Card();
                cartao.encrypted = input.cartaoCriptografado;
            	cartao.holder = new Holder();
            	cartao.store = true;
            	Holder hol = new Holder();
            	hol.name = input.nomeCartao;
            	cartao.holder = hol;
            	pagamento.card = cartao;
            	charg.payment_method = pagamento;
            	b2bRequisicao.charges.add(charg);
        	}
            else if(input.formaPagamento.equalsIgnoreCase('boleto')) {
            	PaymentMethod pagamento = new PaymentMethod();
            	pagamento.type = 'BOLETO';
            	Boleto boleto = new Boleto();
            	boleto.due_date = String.valueOf(Date.today().addDays(2));
            	boleto.instruction_lines = new InstructionLines();
            	InstructionLines intrucoes = new InstructionLines();
            	intrucoes.line_1 = 'Realizar o pagamento em qualquer agência bancária.';
            	boleto.instruction_lines = intrucoes;
            	Holder hol = new Holder();
            	hol.name = contact.Name;
            	hol.tax_id = contact.CPF__c;
            	hol.email = contact.Email;
            	hol.address = endereco.address;
            	boleto.holder = hol;
            	pagamento.boleto = boleto;
            	charg.payment_method = pagamento;
            	b2bRequisicao.charges.add(charg);
        	}
        }
    
        String json = System.JSON.serializePretty(b2bRequisicao);
        json = json.replaceAll('currency_x', 'currency');
        json = json.replaceAll('number_x', 'number');
        
        if(Test.isRunningTest() && account.Name.equalsIgnoreCase('error')) {
        	AuraHandledException newException = new AuraHandledException('isError');
			newException.setMessage('isError');
			throw newException;
        }
        
        return json;
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
        public String reference_id;
        public String description;
        public Amount amount;
        public PaymentMethod payment_method;
    }

    public class Amount {
        public Decimal value;
        public String currency_x;
    }

    public class PaymentMethod {
        public String type;
        public Integer installments;
        public Boolean capture;
        public Card card;
        public Boleto boleto;
    }

    public class Card {
        public String id;
        public String encrypted;
        public String number_x;  
        public String exp_month;    
        public String exp_year;
        public String security_code;
        public Holder holder;
        public Boolean store;
    }

    public class Holder {
        public String name;
        public String tax_id;
        public String email;
        public Address address;
    }

    public class Boleto {
        public String due_date;
        public InstructionLines instruction_lines;
        public Holder holder;
    }

    public class InstructionLines {
        public String line_1;
        public String line_2;
    }

    public class QrCodes {
		public Amount amount;
		public String expiration_date;
	}
}