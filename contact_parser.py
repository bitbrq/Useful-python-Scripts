import os
import base64
from db import initialize_database, insert_contact, insert_record, duplicate_verify

def parse_vcf(file_path):
    with open(file_path, 'r') as file:
        number_id_arr = []
        label_id_arr = []
        birthday_id, name_id, full_name_id, organization_id, address_id, num, email_id = None, None, None, None, None, None, None
        for line in file:
            if line.startswith('BEGIN'):
                begin_vcard = line
            elif line.startswith('BDAY:'):
                birthday_split = line.split('BDAY:')
                if len(birthday_split) > 1:
                    birthday = birthday_split[1].split('\n')[0]
                    birthday_id = insert_record(birthday, 'birthday', 'birthday')

            elif line.startswith('N:'):
                name_split = line.split('N:')
                if len(name_split) > 1:
                    name = name_split[1].split('\n')[0]
                    name_id = insert_record(name, 'name', 'name')

            elif line.startswith('FN:'):
                full_name_split = line.split('FN:')
                if len(full_name_split) > 1:
                    full_name = full_name_split[1].split('\n')[0]
                    full_name_id = insert_record(full_name, 'full_name', 'full_name')

            elif line.startswith('ORG:'):
                organization_split = line.split('ORG:')
                if len(organization_split) > 1:
                    organization = organization_split[1].split('\n')[0]
                    organization_id = insert_record(organization, 'organization', 'organization')

            elif line.startswith('ADR:'):
                address_split = line.split('ADR:')
                if len(address_split) > 1:
                    address = address_split[1].split('\n')[0]
                    address_id = insert_record(address, 'address', 'address')

            elif line.startswith('TEL;'):
                number_splits = line.split('TEL;')
                for number_split in number_splits[1:]:
                    label, number = number_split.split(':')
                    phone_number = ''.join(filter(lambda x: x.isdigit() or x == '+', number.strip()))
                    existing_id = duplicate_verify(phone_number)
                    if existing_id is None:
                        number_id = insert_record(phone_number, 'number', 'number')
                        label_id = insert_record(label.strip(), 'number_label', 'number_label')
                        number_id_arr.append(number_id)
                        label_id_arr.append(label_id)

            elif line.startswith('EMAIL;'):
                email_split = line.split('EMAIL;')
                if len(email_split) > 1:
                    email = email_split[1].split('\n')[0]
                    email_id = insert_record(email, 'email', 'email')

                '''
                elif line.startswith('PHOTO;'):
                photo_split = line.split('PHOTO;')
                if len(photo_split) > 1:
                    # Extract the base64 encoded photo data
                    photo_data_base64 = photo_split[1].split('\n')[0]
                    photo_data = base64.b64decode(photo_data_base64)
                    photo_id = insert_record(photo_data, 'photo', 'photo')
                '''
            elif line.startswith('END'):
                i=0
                for num in number_id_arr:
                    insert_contact(birthday_id, name_id, full_name_id, organization_id, address_id, label_id_arr[i], num, email_id)
                    i=i+1
                i=0
                number_id_arr = []
                label_id_arr = []
                birthday_id, name_id, full_name_id, organization_id, address_id, num, email_id = None, None, None, None, None, None, None


if __name__ == "__main__":
    # Initialize the database
    initialize_database()

    # Path to the VCF file
    vcf_file_path = 'contact.vcf'

    # Check if the VCF file exists
    if os.path.exists(vcf_file_path):
        # Parse the VCF file and add contacts to the database
        parse_vcf(vcf_file_path)
    else:
        print(f"Error: '{vcf_file_path}' does not exist.")
