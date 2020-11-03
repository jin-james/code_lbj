import uuid
from decimal import Decimal
from peewee import Model, ModelSelect
import numpy as np

def field_to_json(value):
    ret = value
    if isinstance(value, list):
        ret = [field_to_json(_) for _ in value]
    elif isinstance(value, dict):
        ret = {k: field_to_json(v) for k, v in value.items()}
    elif isinstance(value, (np.ndarray,)):
        return value.tolist()
    elif isinstance(value, bytes):
        ret = value.decode("utf-8")
    elif isinstance(value, bool):
        ret = int(ret)
    elif isinstance(value, uuid.UUID):
        ret = str(value)
    elif isinstance(value, Decimal):
        ret = float(ret)
    elif isinstance(value, ModelSelect):
        ret = [field_to_json(_) for _ in value]
    elif isinstance(value, Model):
        ret = value.to_json()
    return ret