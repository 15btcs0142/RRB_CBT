import re
from functools import wraps
from flask import request, jsonify

class ValidationError(Exception):
    def __init__(self, message):
        self.message = message
        super().__init__(message)

def validate_field(val, field_name, rules):
    """
    Validate a single field against a rule dictionary.
    Rules:
      - required: bool (default True)
      - type: type (str, int, float, bool)
      - min_length: int
      - max_length: int
      - min_value: int/float
      - max_value: int/float
      - regex: str (regex pattern)
      - allowed_values: set/list/tuple
    """
    is_required = rules.get('required', True)

    if val is None or (isinstance(val, str) and val.strip() == ''):
        if is_required:
            raise ValidationError(f"Field '{field_name}' is required.")
        return val

    target_type = rules.get('type', str)
    typed_val = val

    if target_type is int:
        try:
            typed_val = int(val)
        except (ValueError, TypeError):
            raise ValidationError(f"Field '{field_name}' must be an integer.")
    elif target_type is float:
        try:
            typed_val = float(val)
        except (ValueError, TypeError):
            raise ValidationError(f"Field '{field_name}' must be a number.")
    elif target_type is bool:
        if isinstance(val, bool):
            typed_val = val
        elif isinstance(val, str):
            if val.lower() in ('true', '1', 'yes'):
                typed_val = True
            elif val.lower() in ('false', '0', 'no'):
                typed_val = False
            else:
                raise ValidationError(f"Field '{field_name}' must be a boolean.")
        else:
            raise ValidationError(f"Field '{field_name}' must be a boolean.")
    elif target_type is str:
        if not isinstance(val, str):
            raise ValidationError(f"Field '{field_name}' must be a string.")
        typed_val = val.strip()

    # Length checks for strings
    if target_type is str:
        if 'min_length' in rules and len(typed_val) < rules['min_length']:
            raise ValidationError(f"Field '{field_name}' must be at least {rules['min_length']} characters long.")
        if 'max_length' in rules and len(typed_val) > rules['max_length']:
            raise ValidationError(f"Field '{field_name}' cannot exceed {rules['max_length']} characters.")
        if 'regex' in rules and rules['regex']:
            if not re.match(rules['regex'], typed_val):
                raise ValidationError(f"Field '{field_name}' contains invalid characters or format.")

    # Range checks for numbers
    if target_type in (int, float):
        if 'min_value' in rules and typed_val < rules['min_value']:
            raise ValidationError(f"Field '{field_name}' must be at least {rules['min_value']}.")
        if 'max_value' in rules and typed_val > rules['max_value']:
            raise ValidationError(f"Field '{field_name}' cannot exceed {rules['max_value']}.")

    # Enum checks
    if 'allowed_values' in rules:
        if typed_val not in rules['allowed_values']:
            raise ValidationError(f"Field '{field_name}' must be one of {list(rules['allowed_values'])}.")

    return typed_val

def validate_schema(schema, is_json=False):
    """
    Decorator to validate request payload against a schema dictionary.
    Rejects any non-matching input with HTTP 400 Bad Request.
    """
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if is_json:
                data = request.get_json(silent=True) or {}
            else:
                data = request.form.to_dict()
                if not data and request.args:
                    data = request.args.to_dict()

            validated_data = {}
            for field_name, rules in schema.items():
                raw_val = data.get(field_name)
                try:
                    validated_val = validate_field(raw_val, field_name, rules)
                    validated_data[field_name] = validated_val
                except ValidationError as e:
                    if is_json or request.is_json:
                        return jsonify({'status': 'error', 'message': e.message}), 400
                    return f"Bad Request: {e.message}", 400

            request.validated_data = validated_data
            return f(*args, **kwargs)
        return decorated_function
    return decorator

# Common Schema Definitions
STUDENT_LOGIN_SCHEMA = {
    'name': {'type': str, 'required': True, 'min_length': 1, 'max_length': 100, 'regex': r'^[a-zA-Z0-9\s\.\-_]+$'},
    'student_id': {'type': str, 'required': True, 'min_length': 1, 'max_length': 50, 'regex': r'^[a-zA-Z0-9_\-]+$'},
    'class': {'type': str, 'required': True, 'min_length': 1, 'max_length': 20, 'regex': r'^[a-zA-Z0-9\s\-_]+$'},
    'section': {'type': str, 'required': False, 'max_length': 10, 'regex': r'^[a-zA-Z0-9\s\-_]*$'},
    'subject': {'type': str, 'required': True, 'min_length': 1, 'max_length': 100, 'regex': r'^[a-zA-Z0-9\s\.\-_]+$'},
    'test_no': {'type': str, 'required': False, 'max_length': 50, 'regex': r'^[a-zA-Z0-9\s\-_]*$'}
}

ADMIN_LOGIN_SCHEMA = {
    'username': {'type': str, 'required': True, 'min_length': 1, 'max_length': 50, 'regex': r'^[a-zA-Z0-9_\-\.@]+$'},
    'password': {'type': str, 'required': True, 'min_length': 1, 'max_length': 100}
}

SAVE_ANSWER_SCHEMA = {
    'question_id': {'type': int, 'required': True, 'min_value': 1},
    'selected_option': {'type': str, 'required': False, 'max_length': 5, 'regex': r'^[A-Z]*$'}
}

REATTEMPT_REQUEST_SCHEMA = {
    'reason': {'type': str, 'required': False, 'max_length': 500}
}
