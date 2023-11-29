from BasePBSTransform import BasePBSTransform

class MailDropOverpaymentTransform(BasePBSTransform):
    def __init__(self, input_file_path):
        super().__init__(input_file_path)

    def start_transformation(self):
        input_file_path = self.input_file_path
        print("Starting Transformation for Mail Drop: Overpayment")
