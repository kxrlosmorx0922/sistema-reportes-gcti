from sqlalchemy import Column, Integer, String, Boolean, ForeignKey, DateTime, create_engine
from sqlalchemy.orm import declarative_base, relationship
from sqlalchemy.sql import func

Base = declarative_base()

class Empresa(Base):
    """
    Tabla principal para los clientes corporativos.
    Cada empresa registrada aquí tendrá sus propios colaboradores y usuarios.
    """
    __tablename__ = 'empresas'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(100), nullable=False, unique=True)
    fecha_creacion = Column(DateTime, server_default=func.now())
    
    # Relaciones inversas (Para acceder fácil desde el código)
    usuarios = relationship("Usuario", back_populates="empresa", cascade="all, delete-orphan")
    colaboradores = relationship("Colaborador", back_populates="empresa", cascade="all, delete-orphan")
    categorias_demograficas = relationship("CategoriaDemografica", back_populates="empresa", cascade="all, delete-orphan")


class Usuario(Base):
    """
    Maneja el inicio de sesión y la seguridad (RBAC).
    Determina si es el Coordinador (admin) o un Cliente.
    """
    __tablename__ = 'usuarios'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    email = Column(String(100), nullable=False, unique=True, index=True)
    password_hash = Column(String(255), nullable=False) # Contraseña encriptada
    rol = Column(String(20), nullable=False) # 'admin' (Coordinador) o 'cliente' (Empresa)
    
    # Llave foránea: Si es 'cliente', DEBE estar amarrado a una empresa.
    # Si es 'admin', puede ser Null (porque supervisa todas las empresas).
    empresa_id = Column(Integer, ForeignKey('empresas.id'), nullable=True)
    
    empresa = relationship("Empresa", back_populates="usuarios")


class Colaborador(Base):
    """
    Guarda los datos fijos de la base de datos de colaboradores.
    El E-Mail tiene un INDEX porque es nuestra llave maestra para el XLOOKUP de Python.
    """
    __tablename__ = 'colaboradores'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    identificacion = Column(String(50), nullable=False)
    nombre = Column(String(150), nullable=False)
    email = Column(String(100), nullable=False, index=True) # <-- LLAVE DEL XLOOKUP
    
    empresa_id = Column(Integer, ForeignKey('empresas.id'), nullable=False)
    
    empresa = relationship("Empresa", back_populates="colaboradores")
    valores_demograficos = relationship("ValorDemografico", back_populates="colaborador", cascade="all, delete-orphan")
    participaciones = relationship("Participacion", back_populates="colaborador", cascade="all, delete-orphan")


class CategoriaDemografica(Base):
    """
    Guarda las columnas demográficas detectadas dinámicamente.
    Ejemplo de registros: 
    - id: 1, nombre: 'Lugar de Trabajo', empresa_id: 5
    - id: 2, nombre: 'Área de Desempeño 1', empresa_id: 5
    """
    __tablename__ = 'categorias_demograficas'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    nombre = Column(String(100), nullable=False)
    empresa_id = Column(Integer, ForeignKey('empresas.id'), nullable=False)
    
    empresa = relationship("Empresa", back_populates="categorias_demograficas")
    valores = relationship("ValorDemografico", back_populates="categoria", cascade="all, delete-orphan")


class ValorDemografico(Base):
    """
    Esta es la tabla mágica (EAV). Cruza al colaborador con su categoría y guarda el valor.
    Ejemplo de cómo se vería internamente:
    - colaborador_id: 101 (Carlos), categoria_id: 1 (Lugar de trabajo) -> valor: 'Piedecuesta'
    - colaborador_id: 101 (Carlos), categoria_id: 2 (Área de Desempeño 1) -> valor: 'Gerencia Hospital'
    """
    __tablename__ = 'valores_demograficos'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    colaborador_id = Column(Integer, ForeignKey('colaboradores.id'), nullable=False)
    categoria_id = Column(Integer, ForeignKey('categorias_demograficas.id'), nullable=False)
    valor = Column(String(150), nullable=False)
    
    colaborador = relationship("Colaborador", back_populates="valores_demograficos")
    categoria = relationship("CategoriaDemografica", back_populates="valores")


class Participacion(Base):
    """
    Guarda el histórico de las respuestas que el coordinador va cargando varias veces al día.
    Si hay un registro aquí para el colaborador, significa que YA contestó.
    """
    __tablename__ = 'participaciones'
    
    id = Column(Integer, primary_key=True, autoincrement=True)
    colaborador_id = Column(Integer, ForeignKey('colaboradores.id'), nullable=False)
    fecha_respuesta = Column(DateTime, nullable=True) # Fecha exacta en que contestó según el aplicativo
    contesto = Column(Boolean, default=True)
    fecha_carga = Column(DateTime, server_default=func.now()) # Cuándo subió el coordinador este reporte
    
    colaborador = relationship("Colaborador", back_populates="participaciones")