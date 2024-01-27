class PackageInfo:
    def __init__(self, first_name=None, last_name=None, street=None, house_number=None, city=None, state=None,
                 zipcode=None, case_number=None, product_sku=None, box_name=None, date_received=None, zt01_number=None,
                 vl01n_number=None, notes=None, email=None):
        self.first_name = first_name
        self.last_name = last_name
        self.street = street
        self.house_number = house_number
        self.city = city
        self.state = state
        self.zipcode = zipcode
        self.case_number = case_number
        self.product_sku = product_sku
        self.box_name = box_name
        self.date_received = date_received
        self.zt01_number = zt01_number
        self.vl01n_number = vl01n_number
        self.notes = notes
        self.email = email