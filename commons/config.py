from file_operators.PBSFileOperator import PBSFileOperator
from file_operators.HCMFileOperator import HCMFileOperator

file_operator_classes = {
    "PBS": PBSFileOperator,
    "HCM": HCMFileOperator
}

base_folders = {
    "PBS": "MNL Paygroup/Oncycle"
}