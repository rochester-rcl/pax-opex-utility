from pyPreservica import *

def folder_desc_size(window, mline, user_name, pass_word, ten_ancy, ser_ver, two_factorcb, two_factorkey, folder_uuid):
    mline.update('----GENERATING STORAGE SIZE REPORT FOR FOLDER----\n', append=True, text_color_for_value='white', background_color_for_value='red')
    window.refresh()
    if two_factorcb == True:
        client = EntityAPI(username=user_name, password=pass_word, tenant=ten_ancy, server=ser_ver, two_fa_secret_key=two_factorkey)
    else:
        client = EntityAPI(username=user_name, password=pass_word, tenant=ten_ancy, server=ser_ver)
    folder_target = client.folder(folder_uuid)
    folder_size = 0
    total_assets = 0
    total_files = 0
    for asset in filter(only_assets, client.all_descendants(folder_target.reference)):
        total_assets += 1
        for representation in client.representations(asset):
            for content_object in client.content_objects(representation):
                for generation in client.generations(content_object):
                    for bitstream in generation.bitstreams:
                        total_files += 1
                        folder_size += bitstream.length
                        if (total_files % 2) == 0:
                            mline.update('filename: {} | file size: {}\n'.format(bitstream.filename, bitstream.length), append=True, text_color_for_value='black', background_color_for_value='white')
                            window.refresh()
                        else:
                            mline.update('filename: {} | file size: {}\n'.format(bitstream.filename, bitstream.length), append=True, text_color_for_value='white', background_color_for_value='black')
                            window.refresh()
    folder_gb = round(folder_size / (1024 * 1024 * 1024), 2)
    folder_tb = round(folder_size / (1024 * 1024 * 1024 * 1024), 2)
    mline.update('''Storage Report for Folder
Title: {title}
Ref ID: {ref_id}
Bytes: {bytes}
GB: {gb}
TB: {tb}\n'''.format(title=folder_target.title, ref_id=folder_target.reference, bytes=folder_size, gb=folder_gb, tb=folder_tb), append=True, text_color_for_value='white', background_color_for_value='green')
    window.refresh()