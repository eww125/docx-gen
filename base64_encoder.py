def encode_file(file):
    import base64
    data = open(file, 'rb').read()
    base64_encoded = base64.b64encode(data)
    file_string = str(base64_encoded)[2:]
    file_string = file_string[:-1]
    return file_string