# pdfcraft 0-1-3 19-04-2026 by mrchn

import importlib, os, sys, hashlib, datetime, secrets, subprocess, platform, time, concurrent.futures, dataclasses
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) ; TEMPLATE_PATH = os.path.join(BASE_DIR, 'template.docx')
TOKEN_HEX_LENGTH, COLOR_LOG, COLOR_ERR, COLOR_SUCCESS, COLOR_RESET = 4, '\033[94m', '\033[91m', '\033[92m', '\033[0m'

@dataclasses.dataclass
class Client:
	user_id: int
	username: str = None
	name: str = None
	birth: str = None
	email: str = None
	phone: str = None
	passport: str = None
	departament: str = None
	issue: str = None
	code: str = None
	address: str = None
	contract_id: str = None
	tracklist: str = None
	authors_text: str = None
	authors_music: str = None

class CMD:
	def __init__(self, process=None): self.process = process
	def log(self, info): print(f'{COLOR_LOG}[pdfcraft{f': {self.process}' if self.process else ''}]{COLOR_RESET} {info}')
	def err(self, info): print(f'{COLOR_ERR}[pdfcraft{f': {self.process}' if self.process else ''}]{COLOR_RESET} {info}')
def module_import(module_name, import_name=None):
	''' Динамический импорт модулей (Dynamic modules import) '''
	name, cmd = import_name or module_name, CMD(process='module')
	try: return importlib.import_module(name)
	except ImportError:
		try:
			cmd.log(f'installing {name}...') ; subprocess.check_call([sys.executable, '-m', 'pip', 'install', module_name])
			return importlib.import_module(name)
		except Exception as e: cmd.err(f'{name} not installed: {e}\n')
def get_file_hash(path):
	''' Получение хэша файла (Get file hash) '''
	sha256_hash = hashlib.sha256()
	with open(path, 'rb') as file:
		for byte_block in iter(lambda: file.read(4096), b''): sha256_hash.update(byte_block)
	return sha256_hash.hexdigest()
def generate_signature_hash(user_id, contract_id, date, SECRET_SALT):
	''' Создает уникальный хэш подписи на основе данных договора (Creates a unique signature hash based on the contract data) '''
	data_string = f'{user_id}:{contract_id}:{date}:{SECRET_SALT}' ; return hashlib.sha256(data_string.encode('utf-8')).hexdigest().upper()

class Convert: # res = Сonvert().process('document1.docx', out='document2.pdf', via='docx2pdf') or res = Сonvert().process('doc.docx')
	def __init__(self, remove_in=False, logging=False):
		self.cmd, self.docx2pdf, self.remove_in, self.logging = CMD(process='converter'), None, remove_in, logging

	def process(self, in_path, out=None, via=None, remove_in=None):
		remove_in = remove_in or self.remove_in
		if not via: via = 'docx2pdf' if platform.system() == 'Windows' else 'libre'
		via = via.lower()
		if via in ['word', 'docx2pdf']: return self.via_docx2pdf(str(in_path), out_path=out, remove_in=remove_in)
		else: out_dir = os.path.dirname(str(out)) if out else '.' ; return self.via_libre(str(in_path), out_dir=out_dir, remove_in=remove_in)

	def via_libre(self, in_path, out_dir=None, remove_in=None, pdfa=False):
		'''
		Конвертация через libreoffice (для linux-систем или в отсутствии Word на Винде)
		Convertation via libreoffice (for linux or when Word not installed on Windows)
		'''
		remove_in = remove_in or self.remove_in
		if platform.system() == 'Windows':
			default_path = r'C:\Program Files\LibreOffice\program\soffice.exe'
			proc_path = default_path if os.path.exists(default_path) else 'soffice'
		else: proc_path = 'libreoffice'
		cmdlet = [proc_path, '--headless', '--convert-to', 'pdf:PDF_A_1B' if pdfa is True else 'pdf', '--outdir', str(out_dir), str(in_path)]
		try:
			subprocess.run(cmdlet, check=True)
			filename = os.path.splitext(os.path.basename(in_path))[0] + '.pdf' ; out_path = os.path.join(out_dir, filename)
			if out_path:
				if remove_in and os.path.exists(str(in_path)): os.remove(str(in_path))
				if self.logging: self.cmd.log(f'successfuly converted ({in_path}) -> ({out_path})')
				return str(out_path)
		except subprocess.CalledProcessError as e: self.cmd.err(f'{e}') ; return None

	def via_docx2pdf(self, in_path, out_path=None, remove_in=False):
		'''
		Конвертация через docx2pdf (Microsoft Word, идеально для Windows-систем)
		Convertation via docx2pdf (Microsoft Word, perfectly for Windows)
		'''
		if not self.docx2pdf: self.docx2pdf = module_import('docx2pdf')
		out_path = str(out_path or str(in_path).replace('.docx', '.pdf'))
		try:
			self.docx2pdf.convert(str(in_path), str(out_path))
			if out_path:
				if remove_in and os.path.exists(str(in_path)): os.remove(str(in_path))
				if self.logging: self.cmd.log(f'successfuly converted ({str(in_path)}) -> ({str(out_path)})')
				return str(out_path)
		except Exception as e: self.cmd.err(f'{str(e)}') ; return False

def _generate_worker(data, in_path=None, out_path=None): # for Generate().batch_process()
	return Generate(logging=False).process(data, in_path=in_path, out_path=out_path, batch=True)
class Generate: # res = Generate().process(data, in_path='document1.docx', out_path='document2.pdf') or res = Generate().process(data)
	def __init__(self, remove_in=True, logging=False):
		self.cmd, self.docxtpl, self.logging = CMD(process='generator'), None, logging
		self.converter = Convert(remove_in=remove_in, logging=logging)

	def validate(self, data: dict, in_path=None) -> list:
		'''
		Возвращает список отсутствующих переменных в data, которые требуются в шаблоне. Если список пуст — всё ок
		(Returns a list of missing resources in the data required by the template. If the list is empty, everything is OK)
		'''
		if not self.docxtpl: self.docxtpl = module_import('docxtpl')
		in_path = str(in_path if in_path else TEMPLATE_PATH)
		tpl = self.docxtpl.DocxTemplate(in_path) ; missing_vars = tpl.get_undeclared_template_variables()
		missing = [var for var in missing_vars if var not in data]
		if missing and self.logging: self.cmd.err(f'missing variables: {str(missing)}')
		return missing

	def process(self, data: dict, in_path=None, out_path=None, batch=False, dry=False) -> str | None:
		'''
		Простая генерация pdf с шаблона или под кастомный вход (Easy PDF generation from a template or custom input)
		- data: Словарь с переменными в .docx (Dictionary with variables in .docx)
		- out_path: Имя файла на выходе (Output file name)
		'''
		start_time = time.time()
		in_path = str(in_path if in_path else TEMPLATE_PATH)
		path_docx = f'out_{str(secrets.token_hex(TOKEN_HEX_LENGTH))}_{str(datetime.datetime.now().strftime("%H%M%S"))}.docx'
		target_path = str(out_path or path_docx.replace('.docx', '.pdf'))
		if self.validate(data, in_path): return None
		if dry and self.logging: self.cmd.log(f'dry would generate: {str(target_path)}') ; return str(target_path)
		if self.logging and batch is False: self.cmd.log(f'processing for {str(in_path)}...')
		try:
			tpl = self.docxtpl.DocxTemplate(str(in_path)) ; tpl.render(data) ; tpl.save(str(path_docx))
			res = self.converter.process(str(path_docx), out=str(target_path))
			if res:
				if self.logging and batch is False: self.cmd.log(f'generated in {str(time.time() - start_time:.2f)}s, path: {str(res)}')
				return str(res)
			else: return None
		except Exception as e:
			self.cmd.err(f'error: {e}')
			if os.path.exists(path_docx): os.remove(path_docx)
			return None

	def batch_process(self, data_list: list[dict], out_dir: str = 'out', in_path=None, prefix=None, max_workers=4) -> list[str]:
		'''
		Массовая генерация документов из списка словарей (Batch generation of documents from a list of dictionaries)
		data_list: Список словарей с данными для каждого документа (List of dictionaries with data for each document)
		out_dir: Директория на выход (Output directory)
		'''
		start_time = time.time() ; os.makedirs(out_dir, exist_ok=True) ; results = []
		tasks = [data for data in data_list if not self.validate(data, in_path=in_path)]
		if self.logging: self.cmd.log(f'starting batch process (dir: {str(out_dir)}, workers: {max_workers}) for {len(data_list)} documents...')
		with concurrent.futures.ProcessPoolExecutor(max_workers=max_workers) as executor:
			futures = []
			for index, data in enumerate(tasks):
				doc_id = data.get('id', f'doc_{index}')
				target_path = os.path.join(out_dir, f'{str(prefix)}_{doc_id}.pdf' if prefix else f'out_{doc_id}.pdf')
				futures.append(executor.submit(_generate_worker, data, in_path=str(in_path), out_path=target_path))
			for future in futures:
				res = future.result()
				if res: results.append(res)
		if self.logging: self.cmd.log(f'batch process finished in {time.time() - start_time:.2f}s, {len(results)}/{len(data_list)} generated.')
		return results

class Contracts:
	def __init__(self, logging=True, prefix=None, SECRET_SALT=None, template=TEMPLATE_PATH):
		self.generator, self.contract_prefix, self.SECRET_SALT, self.template_path = Generate(logging=logging), prefix, SECRET_SALT, template

	def generate_contract_id(self, token_hex_len=TOKEN_HEX_LENGTH):
		return f'{datetime.datetime.now().strftime("%y")}{secrets.token_hex(int(token_hex_len)).upper()}'

	def generate(
		self, client, SECRET_SALT=None, signed=False, user_id=None, in_path=None, out_path=None, prefix=None, token_hex_len=TOKEN_HEX_LENGTH
		) -> str | None:
		'''
		Generation license contract in telegram-bot (Takes data from Client-class objects)
		Генерация лицензионного договора в телеграм-боте (Получает данные объектов класса Client)
		'''
		date, SECRET_SALT = datetime.datetime.now().strftime('%d.%m.%Y'), SECRET_SALT or self.SECRET_SALT
		contract_prefix, contract_id = prefix or self.contract_prefix, self.generate_contract_id(token_hex_len=token_hex_len)
		in_path = str(in_path if in_path else self.template_path)
		out_path = str(out_path if out_path else f'{str(contract_prefix) if contract_prefix else ''}{str(contract_id)}.pdf')
		all_data = {
			# данные клиента
			'client_name': str(client.name or 'Иванов Иван Иванович'),
			'client_birth': str(client.birth or '01.01.2000'),
			'client_passport': str(client.passport or '1234 567890'),
			'client_departament': str(client.departament or 'УМВД России по Смоленской области'),
			'client_issue': str(client.issue or '01.01.2000'),
			'client_code': str(client.code or '670-001'),
			'client_address': str(client.address or 'Смоленск, пр-т. Гагарина, 1'),
			'client_email': str(client.email or 'example@email.com'),
			'client_phone': str(client.phone or '+7 (912) 345-67-89'),
			'client_track': str(getattr(client, 'tracklist', 'трек - исполнитель')),
			'client_text': str(getattr(client, 'authors_text', str(client.name))),
			'client_music': str(getattr(client, 'authors_music', str(client.name))),
			'client_phonogram': f'{getattr(client, 'authors_text', str(client.name))}, {getattr(client, 'authors_music', str(client.name))}',
			# данные для подписи
			'signed': signed, # если False, подпись не поставится ("ПРЕДНАЗНАЧЕН ДЛЯ ОЗНАКОМЛЕНИЯ")
			'signature_date': str(datetime.datetime.now().isoformat()),
			'contract_id': str(contract_id), 'date': str(date),
			'contract_id_prefix': str(f'{contract_prefix}-') if contract_prefix is not None else '',
			'signature_user_id': str(user_id), # id пользователя в мессенджере
			'signature_hash': generate_signature_hash(str(user_id), str(contract_id), str(date), str(SECRET_SALT))
		}
		res = self.generator.process(all_data, in_path=str(in_path), out_path=str(out_path)) ; return str(res)