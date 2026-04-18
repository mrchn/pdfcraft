# pdfcraft 0-1-0 18-04-2026 by mrchn

TEMPLATE_PATH = 'template.docx'
COLOR_LOG, COLOR_ERR, COLOR_SUCCESS, COLOR_RESET = '\033[94m', '\033[91m', '\033[92m', '\033[0m'

import importlib, os, sys, hashlib, datetime, secrets, subprocess, platform, time, concurrent.futures

class CMD:
	def __init__(self, process=None): self.process = process
	def log(self, info): print(f'{COLOR_LOG}[pdfcraft{f' : {self.process}' if self.process else ''}]{COLOR_RESET} {info}')
	def err(self, info): print(f'{COLOR_ERR}[pdfcraft{f' : {self.process}' if self.process else ''}]{COLOR_RESET} {info}')
	class Bot:
		def __init__(self, username): self.username = username
		def log(self, info): print(f'{COLOR_LOG}(@{self.username}){COLOR_RESET} {info}')
		def err(self, info): print(f'{COLOR_ERR}(@{self.username}){COLOR_RESET} {info}')
def module_import(module_name, import_name=None): # import or install pip-module
	name, cmd = import_name or module_name, CMD(process='module')
	try: return importlib.import_module(name)
	except ImportError:
		try:
			cmd.log(f'installing {name}...') ; subprocess.check_call([sys.executable, '-m', 'pip', 'install', module_name])
			return importlib.import_module(name)
		except Exception as e: cmd.err(f'{name} not installed: {e}\n')

def get_file_hash(path):
	sha256_hash = hashlib.sha256()
	with open(path, 'rb') as file:
		for byte_block in iter(lambda: file.read(4096), b''): sha256_hash.update(byte_block)
	return sha256_hash.hexdigest()
def gen_contract_id(prefix=None, token_hex_len=2):
	return f'{datetime.datetime.now().strftime("%y")}{secrets.token_hex(int(token_hex_len)).upper()}'

class Сonvert: # res = Сonvert().process('document1.docx', out='document2.pdf', via='docx2pdf') or res = Сonvert().process('doc.docx')
	def __init__(self, remove_in=False, logging=False):
		self.cmd, self.docx2pdf, self.remove_in, self.logging = CMD(process='convert'), None, remove_in, logging

	def process(self, in_path, out=None, via=None, remove_in=None):
		remove_in = remove_in or self.remove_in
		if not via: via = 'docx2pdf' if platform.system() == 'Windows' else 'libre'
		via = via.lower()
		if via in ['word', 'docx2pdf']: return self.via_docx2pdf(str(in_path), out_path=out, remove_in=remove_in)
		else: out_dir = os.path.dirname(out) if out else '.' ; return self.via_libre(str(in_path), out_dir=out_dir, remove_in=remove_in)

	def via_libre(self, in_path, out_dir=None, remove_in=None, pdfa=False):
		'''
		Конвертация через libreoffice (для linux-систем или в отсутствии Word на Винде)
		'''
		remove_in = remove_in or self.remove_in
		proc_path = r'C:\Program Files\LibreOffice\program\soffice.exe' if platform.system() == 'Windows' else 'libreoffice'
		cmdlet = [proc_path, '--headless', '--convert-to', 'pdf:PDF_A_1B' if pdfa is True else 'pdf', '--outdir', str(out_dir), str(in_path)]
		try:
			subprocess.run(cmdlet, check=True)
			filename = os.path.splitext(os.path.basename(in_path))[0] + '.pdf'
			out_path = os.path.join(out_dir, filename)
			if out_path:
				if remove_in: os.remove(str(in_path))
				if self.logging: self.cmd.log(f'successfuly converted ({in_path}) -> ({out_path})')
				return out_path
		except subprocess.CalledProcessError as e: self.cmd.err(f'{e}') ; return None

	def via_docx2pdf(self, in_path, out_path=None, remove_in=False):
		'''
		Конвертация через docx2pdf (Microsoft Word, идеально для Windows-систем)
		'''
		if not self.docx2pdf: self.docx2pdf = module_import('docx2pdf')
		out_path = str(out_path) or str(in_path).replace('.docx', '.pdf')
		try:
			self.docx2pdf.convert(str(in_path), out_path)
			if out_path:
				if remove_in: os.remove(str(in_path))
				if self.logging: self.cmd.log(f'successfuly converted ({in_path}) -> ({out_path})')
				return out_path
		except Exception as e: self.cmd.err(f'{e}') ; return False

def _generate_worker(data, in_path=None, out_path=None):
	return Generate(logging=False).process(data, in_path=in_path, out_path=out_path, batch=True)
class Generate: # res = Generate().process(data, in_path='document1.docx', out_path='document2.pdf') or res = Generate().process(data)
	def __init__(self, remove_in=True, logging=False):
		self.cmd, self.docxtpl, self.logging = CMD(process='generate'), None, logging
		self.converter = Convert(remove_in=remove_in, logging=logging)

	def validate(self, data: dict, in_path=None) -> list:
		'''
		Возвращает список отсутствующих переменных в data,
		которые требуются в шаблоне. Если список пуст — всё ок.
		'''
		if not self.docxtpl: self.docxtpl = module_import('docxtpl')
		in_path = str(in_path) or TEMPLATE_PATH
		tpl = self.docxtpl.DocxTemplate(in_path) ; missing_vars = tpl.get_undeclared_template_variables()
		missing = [var for var in missing_vars if var not in data]
		if missing and self.logging: self.cmd.err(f'missing variables: {missing}')
		return missing

	def process(self, data: dict, in_path=None, out_path=None, batch=False, dry=False) -> str | None:
		'''
		Простая генерация pdf с шаблона или под кастомный вход.
		data: словарь с переменными в .docx
		out_path: имя файла на выходе.
		'''
		start_time = time.time()
		if self.validate(data, in_path): return None
		in_path = str(in_path) or TEMPLATE_PATH
		path_docx = f'temp_{str(datetime.datetime.now().strftime("%Y%m%d_%H%M%S"))}.docx'
		target_path = out_path or path_docx.replace('.docx', '.pdf')
		if dry and self.logging: self.cmd.log(f'dry would generate: {target_path}') ; return target_path
		if self.logging and batch is False: self.cmd.log(f'processing for {in_path}...')
		try:
			tpl = self.docxtpl.DocxTemplate(in_path) ; tpl.render(data) ; tpl.save(path_docx)
			res = self.converter.process(path_docx, out=target_path)
			if res:
				if self.logging and batch is False: self.cmd.log(f'generated successfuly in {time.time() - start_time:.2f}s, path: {res}')
				return res
			else: return None
		except Exception as e:
			self.cmd.err(f'error: {e}')
			if os.path.exists(path_docx): os.remove(path_docx)
			return None

	def batch_process(self, data_list: list[dict], out_dir: str = 'out', in_path=None, prefix=None, max_workers=4) -> list[str]:
		'''
		Массовая генерация документов из списка словарей.
		data_list: список словарей с данными для каждого документа.
		out_dir: директория на выход.
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