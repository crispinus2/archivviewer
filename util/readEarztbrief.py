import sys, email

if __name__ == '__main__':

    with open(sys.argv[1], 'rb') as infile:
        eml = email.message_from_bytes(infile.read())
                     
        for part in eml.get_payload():
            fnam = part.get_filename()
            if fnam is not None:
                partcont = part.get_payload(decode=True)
                with open(fnam, 'wb') as outfile:
                    outfile.write(partcont)
