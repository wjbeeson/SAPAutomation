from package_info import PackageInfo
import receive_form
import repair_form

def update_package_info(original: PackageInfo, new: PackageInfo, override=False):
    if override:
        for field in original.__dict__:
            if new.__dict__[field] is not None:
                original.__dict__[field] = new.__dict__[field]
    else:
        for field in original.__dict__:
            if original.__dict__[field] is None:
                original.__dict__[field] = new.__dict__[field]
    return original

def switch_to_form(self, form_name):
    form_dict = {
        "RECEIVE": receive_form.ReceiveForm,
        "REPAIR": repair_form.RepairForm,
    }
    driver = self.driver
    package_info = self.get_package_info()
    self.base.destroy()
    form_dict[form_name](driver, package_info)