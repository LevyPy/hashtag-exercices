from flask import render_template, redirect, url_for, flash, request
from site_exemplo import app, database, bcrypt
from site_exemplo.forms import FormLogin, FormCriarConta, FormEditarPerfil, FormCriarPost, FormLimparBase
from site_exemplo.models import Usuario, Post
from flask_login import login_user, logout_user, current_user, login_required
import secrets
import os
from PIL import Image

lista_usuarios = ['Lira', 'João', 'Alon', 'Alessandra', 'Amanda']


@app.route('/')
def home():
    posts = Post.query.order_by(Post.id.desc())
    return render_template('home.html', posts=posts)


@app.route('/contato')
def contato():
    return render_template('contato.html')


@app.route('/usuarios')
@login_required
def usuarios():
    lista_usuarios = Usuario.query.all()
    return render_template('usuarios.html', lista_usuarios=lista_usuarios)


@app.route('/login', methods=['GET', 'POST'])
def login():
    form_login = FormLogin()
    form_criarconta = FormCriarConta()
    if form_login.validate_on_submit() and 'botao_submit_login' in request.form:
        usuario = Usuario.query.filter_by(email=form_login.email.data).first()
        if usuario and bcrypt.check_password_hash(usuario.senha, form_login.senha.data):
            login_user(usuario, remember=form_login.lembrar_dados.data)
            flash(f'Login feito com sucesso no e-mail: {form_login.email.data}', 'alert-success')
            par_next = request.args.get('next')
            if par_next:
                return redirect(par_next)
            else:
                return redirect(url_for('home'))
        else:
            flash(f'Falha no Login. E-mail ou Senha Incorretos', 'alert-danger')
    if form_criarconta.validate_on_submit() and 'botao_submit_criarconta' in request.form:
        senha_cript = bcrypt.generate_password_hash(form_criarconta.senha.data)
        usuario = Usuario(username=form_criarconta.username.data, email=form_criarconta.email.data, senha=senha_cript)
        database.session.add(usuario)
        database.session.commit()
        flash(f'Conta criada para o e-mail: {form_criarconta.email.data}', 'alert-success')
        return redirect(url_for('home'))
    return render_template('login.html', form_login=form_login, form_criarconta=form_criarconta)


@app.route('/sair')
@login_required
def sair():
    logout_user()
    flash(f'Logout Feito com Sucesso', 'alert-success')
    return redirect(url_for('home'))


@app.route('/perfil')
@login_required
def perfil():
    foto_perfil = url_for('static', filename='fotos_perfil/{}'.format(current_user.foto_perfil))
    return render_template('perfil.html', foto_perfil=foto_perfil)


@app.route('/post/criar', methods=['GET', 'POST'])
@login_required
def criar_post():
    form = FormCriarPost()
    if form.validate_on_submit():
        post = Post(titulo=form.titulo.data, corpo=form.corpo.data, autor=current_user)
        database.session.add(post)
        database.session.commit()
        flash('Post Criado com Sucesso', 'alert-success')
        return redirect(url_for('home'))
    return render_template('criarpost.html', form=form)


def salvar_imagem(imagem):
    codigo = secrets.token_hex(8)
    nome, extensao = os.path.splitext(imagem.filename)
    nome_arquivo = nome + codigo + extensao
    caminho_completo = os.path.join(app.root_path, 'static/fotos_perfil', nome_arquivo)
    tamanho = (400, 400)
    imagem_reduzida = Image.open(imagem)
    imagem_reduzida.thumbnail(tamanho)
    imagem_reduzida.save(caminho_completo)
    return nome_arquivo


def atualizar_cursos(form):
    lista_cursos = []
    for campo in form:
        if 'curso_' in campo.name:
            if campo.data:
                lista_cursos.append(campo.label.text)
    return ';'.join(lista_cursos)


@app.route('/perfil/editar', methods=['GET', 'POST'])
@login_required
def editar_perfil():
    form = FormEditarPerfil()
    if form.validate_on_submit():
        current_user.email = form.email.data
        current_user.username = form.username.data
        if form.foto_perfil.data:
            nome_imagem = salvar_imagem(form.foto_perfil.data)
            current_user.foto_perfil = nome_imagem
        current_user.cursos = atualizar_cursos(form)
        database.session.commit()
        flash('Perfil atualizado com Sucesso', 'alert-success')
        return redirect(url_for('perfil'))
    elif request.method == "GET":
        form.email.data = current_user.email
        form.username.data = current_user.username
    foto_perfil = url_for('static', filename='fotos_perfil/{}'.format(current_user.foto_perfil))
    return render_template('editarperfil.html', foto_perfil=foto_perfil, form=form)


@app.route('/post/<post_id>', methods=['GET', 'POST'])
@login_required
def exibir_post(post_id):
    post = Post.query.get(post_id)
    if current_user == post.autor:
        form = FormCriarPost()
        if request.method == 'GET':
            form.titulo.data = post.titulo
            form.corpo.data = post.corpo
        elif form.validate_on_submit():
            post.titulo = form.titulo.data
            post.corpo = form.corpo.data
            database.session.commit()
            flash('Post Atualizado com Sucesso', 'alert-success')
            return redirect(url_for('home'))
    else:
        form = None
    return render_template('post.html', post=post, form=form)


@app.route('/post/<post_id>/excluir', methods=['GET', 'POST'])
@login_required
def excluir_post(post_id):
    post = Post.query.get(post_id)
    if current_user == post.autor:
        database.session.delete(post)
        database.session.commit()
        flash('Post Excluído com Sucesso', 'alert-danger')
        return redirect(url_for('home'))
    else:
        abort(403)

@app.route('/teste/limpando', methods=['GET', 'POST'])
def limpar():
    form = FormLimparBase()
    if form.validate_on_submit():
        if form.file_path.data:
            base_bruta = import_db(form.file_path.data, separator) #caminho do arquivo bruto
            if coluna_stonecode != '':
                sc_doc_col = coluna_stonecode
                sc_doc_key = 'stonecode'
                sql = '[NR_STONECODE]'

            elif coluna_documento != '':
                sc_doc_col = coluna_documento
                sc_doc_key = 'documento'
                sql = '[CNPJ_CPF]'

            else:
                print('Favor preencher qual a coluna do Stonecode ou do Documento.')

            base_nao_tratada = base_bruta
            base_nao_tratada[sc_doc_col] = base_nao_tratada[sc_doc_col].str.replace(' ', '')
            base_nao_tratada[sc_doc_col] = base_nao_tratada[sc_doc_col].str.lstrip('0')
            base_nao_tratada['duplicated'] = base_nao_tratada.duplicated(subset=[sc_doc_col], keep='first')
            lista_sc_doc_non_dup = base_nao_tratada[base_nao_tratada['duplicated'] == False]
            sc_doc_1 = lista_sc_doc_non_dup[sc_doc_col].tolist()
            aux = 0
            total_iterations = math.ceil(len(sc_doc_1) / step)


    return render_template('limpabase.html', form=form)


def import_db(file_path: str, separator: str) -> 'pandas.core.frame.DataFrame':
    if '.xlsx' in file_path[-5:]:
        print('Identificado arquivo Excel, executando rotina de Excel.')
        df_import = pd.read_excel(file_path, dtype=str)

    elif '.csv' in file_path[-4:]:
        print('Identificado arquivo csv, executando rotina de csv.')
        df_import = pd.read_csv(file_path, sep=separator, dtype=str, skipinitialspace=True, skip_blank_lines=True)

    else:
        print('Não é um formato de arquivo compativel, colocar um arquivo em Excel (.xlsx) ou csv (.csv).')

    df_import = df_import.fillna('')
    display(df_import.columns)
    display(df_import.shape)
    return df_import

