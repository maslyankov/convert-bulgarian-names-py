import logging
from json import dump
from csv import writer
from os import makedirs
from time import strftime, sleep

from transliterate import translit, get_available_language_codes
from transliterate.exceptions import LanguagePackNotFound
import xlrd
from ldap3 import Server, Connection

from config import LDAP_SERVER, LDAP_USER_BASE, LDAP_USER, LDAP_PASS, EXCEL_DOCUMENT, NEW_USERNAMES_PREFIX


def get_user_info(first_name=None, last_name=None):
    logging.info("Conecting via LDAP...")

    server = Server(LDAP_SERVER, port=389)
    logging.info(f"Searching with base {LDAP_USER_BASE}")
    fltrs = f"(displayName={first_name})(sn={last_name})" if first_name and last_name else None
    searchParameters = {
        'search_base': LDAP_USER_BASE,
        'search_filter': f'(&(objectclass=person){fltrs if fltrs else ""})',
        'attributes': ['cn', 'givenName', 'sn', 'displayName', 'memberOf']
    }

    with Connection(server, user=LDAP_USER, password=LDAP_PASS) as conn:
        conn.search(**searchParameters)

        return conn.entries[0] if len(conn.entries) == 1 else conn.entries if len(conn.entries) > 1 else None, conn.result


def iterate_excel_file(check_for_duplicates=True):
    logging.info(f'Opening excel file')

    book = xlrd.open_workbook(EXCEL_DOCUMENT)

    sh = book.sheet_by_index(0)

    names = dict()
    names_list = list()

    # iterate through excel
    for row in range(sh.nrows):
        logging.debug(f"Row {row} data:")

        if row == 0:
            continue

        first_name = last_name = full_name = None

        for col in range(sh.ncols):
            cell_val = sh.cell_value(row, col)

            if isinstance(cell_val, float):
                cell_val = int(cell_val)

            if col == 0:
                # logging.debug(f"{str(cell_val)}>", end=" ")
                pass

            if col == 1:  # Which to convert
                try:
                    name_translated = translit(cell_val, 'bg', reversed=True).split()
                    first_name = name_translated[0].capitalize().strip().replace("ь", "y")
                    last_name = name_translated[1].capitalize().strip().replace("ь", "y")
                    full_name = f"{first_name} {last_name}"

                    # logging.debug(f"{full_name} |", end=" ")
                except LanguagePackNotFound:
                    logging.error("Language not found! You can choose from these languages:")
                    logging.info(get_available_language_codes())
                continue

            if col == 2:
                user_id_num = cell_val
                # logging.debug(uid, end="")

                # at last useful col...
                names[full_name] = {
                    'user_id_num': user_id_num,
                    "full_name": full_name,
                    "first_name": first_name,
                    "last_name": last_name
                }
                logging.info(str(names[full_name]))

                if check_for_duplicates and full_name in names_list:
                    logging.error(f"Name {full_name} is already in the list")
                    logging.info("Press enter to continue")
                    input()

                names_list.append(full_name)
    return names, names_list


def save_data(data, filename, folder):
    if isinstance(data, dict):
        with open(f'{folder}/{filename}.json', 'w') as outfile:
            dump(data, outfile)
    elif isinstance(data, list):
        with open(f'{folder}/{filename}.csv', 'w') as outfile:
            wr = writer(outfile)
            wr.writerows(data)


def parse():
    # Create data folder
    subfolder = 'data/' + strftime("%Y%m%d-%H%M%S")
    makedirs(subfolder, exist_ok=True)

    # Logging
    rootLogger = logging.getLogger()
    rootLogger.setLevel(logging.DEBUG)

    logFormatter = logging.Formatter("%(asctime)s [%(funcName)s] [%(levelname)-5.5s]  %(message)s")

    fileLogsHandler = logging.FileHandler(f"{subfolder}/output.log")
    fileLogsHandler.setFormatter(logFormatter)
    rootLogger.addHandler(fileLogsHandler)

    consoleHandler = logging.StreamHandler()
    consoleHandler.setFormatter(logFormatter)
    rootLogger.addHandler(consoleHandler)

    # Parse Excel File with data (with cyrilic names)
    names, names_list = iterate_excel_file()
    save_data(names, "excel_names", subfolder)

    # Get users data from LDAP
    users_data, _ = get_user_info()

    # Start doing matching
    found_legacy = dict()
    found_new = dict()

    errored = dict()
    skipped = dict()

    for user in users_data:
        logging.info(f"================= NEXT USER =================")
        logging.info(f"user={user}")
        logging.info(f"uid {user.cn}, {user.displayName}")
        if "systemaccounts" in str(user.memberOf):
            logging.info("User seems to be sys user. Skipping.")
            skipped[str(user.cn)] = {
                "cn": str(user.cn),
                "displayName": str(user.displayName),
                "memberOf": str(user.memberOf),
                "givenName": str(user.givenName),
                "sn": str(user.sn)
            }
            continue

        if str(user.cn).startswith(NEW_USERNAMES_PREFIX):
            found_new[str(user.cn)] = {
                "cn": str(user.cn),
                "displayName": str(user.displayName),
                "givenName": str(user.givenName),
                "sn": str(user.sn),
                "user_id_num": int(str(user.cn).strip(NEW_USERNAMES_PREFIX))
            }
            continue

        # logging.info(f"uid {user['uid']}, {user['full_name']}")
        fullname = f"{user.givenName} {user.sn}"
        if fullname not in names_list:
            logging.error(f"User displayName {fullname} not found in names_list")

            frst_name = str(user.givenName).replace("ia","iya")
            fullname = f"{frst_name} {user.sn}"

            found_bool = False

            logging.error(f"Retrying with {fullname}")
            if fullname in names_list:
                user_uid = str(user.cn)
                logging.info(f"Found username -> {user_uid}")
                logging.info(f"Found user id -> {names[fullname]['user_id_num']}")
                found_legacy[user_uid] = names[fullname]
                found_legacy[user_uid]['uid'] = user_uid
                logging.info("-------------------------------------")
            else:
                logging.error("Tried to retry but failed...")

                logging.info("Trying only with surname...")
                for f_name in names_list:
                    firstname = str(user.givenName)
                    if str(user.sn) in f_name:
                        if firstname[0:2] not in f_name:
                            continue

                        user_uid = str(user.cn)
                        logging.info(f"Found username -> {user_uid}")
                        found_legacy[user_uid] = names[f_name]
                        found_legacy[user_uid]['uid'] = user_uid
                        found_bool = True
                        logging.info("-------------------------------------")
                        break

                if found_bool:
                    continue
        else:
            user_uid = str(user.cn)
            logging.info(f"Found username -> {user_uid}")
            found_legacy[user_uid] = names[fullname]
            found_legacy[user_uid]['uid'] = user_uid
            logging.info("-------------------------------------")
            continue

        logging.info("All attempts were unsuccessful!")
        errored[fullname] = {
            "cn": str(user.cn),
            "displayName": str(user.displayName),
            "memberOf": str(user.memberOf),
            "givenName": str(user.givenName),
            "sn": str(user.sn)
        }
        logging.info("-------------------------------------")


    logging.info('ERRORED:')
    logging.info(errored.keys())

    logging.info('----------- RECAP -----------')

    logging.info(f"Errored names: {len(errored)}")

    logging.info(f"found legacy usernames: {len(found_legacy)}/{len(names)-len(found_new)}")
    logging.info(f"found new usernames: {len(found_new)}/{len(names)-len(found_legacy)}")

    logging.info(f"(missing {len(names) - len(found_legacy) - len(found_new)})")
    logging.info(f"Skipped users: {len(skipped)}")

    save_data(found_legacy, "found_legacy", subfolder)
    save_data(found_new, "found_new", subfolder)

    save_data(errored, 'errored_names', subfolder)
    save_data(skipped, 'skipped', subfolder)


if __name__ == '__main__':
    parse()
