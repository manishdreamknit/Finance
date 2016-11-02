
def apply_customized_operation(value, operation):
	if operation == 'make_lower':
		return value.lower()

	elif operation == 'make_upper':
		return value.upper()

	elif operation == 'housing_status':
		value = value.upper()
		if value == 'OWNOUTRIGHT':
			return value[0:3] + ' ' + value[3:]
		elif value == 'MORTGAGE':
			return 'HOMEOWNER'

	elif operation == 'get_years':
		return int(value) / 12

	elif operation == 'get_months':
		return int(value) % 12

	elif operation == 'dob':
		return value[5:7] + value[8:10] + value[0:4]

	elif operation == 'street_number':
		# Sending some default value
		return '556'

	elif operation == 'street_name':
		return 'SE 01'

	elif operation == 'street_type':
		return 'ST'

	elif operation == 'finance_method':
		if value == 'Lease':
			return 'lease'
		return 'buy'

	elif operation == 'trade_in':
		if value == []:
			return 'TradeNo'
		else:
			return 'TradeYes'

	elif operation == 'vehicle_in':
		if value == []:
			return 'VehicleNo'
		else:
			return 'VehicleYes'

	elif operation == 'purchase_type':
		return value.lower() + 'Veh'

	elif operation == 'coapp_in':
		if value == []:
			return 'individually'
		return 'joint'

	elif operation == 'relationship':
		value = value.lower()
		if value not in ['spouse', 'relative']:
			return 'other'
		return value

	return value


def get_current_element_attribute(element, attrrib_name):
	return element.get(attrrib_name, False)
