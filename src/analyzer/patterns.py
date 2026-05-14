"""
Padrões, palavras-chave e expressões regulares para análise de documentos
de projetos técnicos em português brasileiro (cenografia, produção, orçamento).
"""
import re

# ── Medidas ───────────────────────────────────────────────────────────────────
# Cada tupla: (padrão regex, tipo, lista_de_grupos_para_campos)
# grupos: None = pegar m.group(0); lista = índices dos grupos
MEASURE_PATTERNS_LIST = [
    # 3D: 2,50 x 1,80 x 0,90 m/cm/mm
    (r'(\d+[,.]?\d*)\s*[xX×]\s*(\d+[,.]?\d*)\s*[xX×]\s*(\d+[,.]?\d*)\s*(cm|mm|m\b|metros?)',
     'LxAxP', ['largura', 'altura', 'profundidade']),
    # 2D: 2,50 x 1,80 m/cm/mm
    (r'(\d+[,.]?\d*)\s*[xX×]\s*(\d+[,.]?\d*)\s*(cm|mm|m\b|metros?)',
     'LxA', ['largura', 'altura']),
    # Área: 12,5 m² ou 12.5m2
    (r'(\d+[,.]?\d*)\s*m[²2]',
     'area', ['area']),
    # Metro linear: 10ml ou 10 m.l. ou 10 metros lineares
    (r'(\d+[,.]?\d*)\s*(?:m\.?l\.?|metros?\s+lineares?)',
     'ml', ['metro_linear']),
    # Diâmetro: Ø50cm ou ø 1,20m
    (r'[ØøOo∅]\s*(\d+[,.]?\d*)\s*(cm|mm|m\b)',
     'diametro', ['comprimento']),
    # Espessura: e=15mm ou esp. 18mm ou 15mm de espessura
    (r'(?:e\s*=\s*|esp\.?\s*)(\d+[,.]?\d*)\s*(cm|mm|m\b)',
     'espessura', ['espessura']),
    (r'(\d+[,.]?\d*)\s*(cm|mm|m\b)\s+de\s+espessura',
     'espessura', ['espessura']),
    # Simples: 2,50m ou 150cm ou 1500mm
    (r'(\d+[,.]?\d*)\s*(cm|mm|m\b)',
     'simples', ['comprimento']),
]

# ── Categorias de itens ───────────────────────────────────────────────────────
CATEGORY_MAP = {
    'Cenografia': [
        'cenografia', 'cenográfico', 'cenografic', 'backdrop', 'backwall', 'backrop',
        'painel de fundo', 'parede cen', 'estrutura cen', 'display', 'stand', 'estande',
        'expositor', 'portal', 'arco', 'torre', 'totem cenográfico', 'módulo',
    ],
    'Comunicação Visual': [
        'comunicação visual', 'cv ', ' cv', 'banner', 'totem', 'adesivo', 'adesivação',
        'lona', 'impressão', 'gráfico', 'arte final', 'facing', 'placa', 'letreiro',
        'letra caixa', 'letra em relevo', 'sinalização', 'sinalética', 'plotagem',
        'front light', 'back light', 'luminoso', 'acm', 'fachada', 'painel gráfico',
        'roll-up', 'rollup', 'x-banner', 'display gráfico',
    ],
    'Mobiliário': [
        'mesa', 'cadeira', 'sofá', 'sofa', 'banco', 'puff', 'poltrona', 'lounge',
        'bancada', 'balcão', 'balcao', 'estante', 'armário', 'armario', 'rack',
        'aparador', 'criado', 'buffet', 'vitrine', 'prateleira', 'nicho',
        'mobiliário', 'mobili', 'furniture', 'mobília', 'cama', 'escrivaninha',
    ],
    'Estrutura': [
        'estrutura', 'metalon', 'tubular', 'treliça', 'trelica', 'perfil',
        'viga', 'coluna', 'pilar', 'suporte', 'sapata', 'base metálica',
        'chapas', 'chapa', 'ferro', 'aço', 'aco', 'alumínio', 'aluminio',
        'box truss', 'boxtruss', 'grid', 'grelha', 'andaime', 'mezzanino',
    ],
    'Acabamento': [
        'acabamento', 'revestimento', 'pintura', 'verniz', 'lacagem', 'laca',
        'lixa', 'stain', 'textura', 'texturizado', 'emassamento', 'massa',
        'piso', 'carpete', 'carpet', 'paginação', 'rejunte', 'pastilha',
        'porcelana', 'cerâmica', 'ceramica', 'pedra', 'mármore', 'marmore',
        'granito', 'laminado', 'formica',
    ],
    'Objeto Decorativo': [
        'objeto', 'decoração', 'decorativo', 'adorno', 'peça decorativa',
        'arranjo', 'vaso', 'cachepot', 'escultura', 'quadro', 'tela',
        'espelho', 'luminária decorativa', 'lustre', 'pendente', 'abajur',
        'tapete', 'almofada', 'cortina', 'persiana', 'trilho',
    ],
    'Paisagismo': [
        'paisagismo', 'planta', 'vegetal', 'jardim', 'grama', 'relva',
        'flores', 'floral', 'arbust', 'árvore', 'arvore', 'substrato',
        'pedriscos', 'samambaia', 'suculenta', 'cactos', 'orquídea',
        'folhagem', 'musgo', 'bambu', 'graveto', 'galho',
    ],
    'Iluminação': [
        'iluminação', 'iluminacao', 'luminária', 'luminaria', 'spot',
        'refletor', 'projetor', 'led', 'fita led', 'neon', 'neon flex',
        'par 64', 'par64', 'moving head', 'moving', 'beam', 'wash',
        'lâmpada', 'lampada', 'bulbo', 'driver', 'dmx', 'dimmer',
        'perfil led', 'pendente', 'cob', 'luz', 'ponto de luz',
    ],
    'Elétrica': [
        'elétrica', 'eletrica', 'quadro elétrico', 'disjuntor', 'tomada',
        'cabo', 'fio', 'eletroduto', 'conduite', 'plug', 'extensão',
        'régua', 'no-break', 'gerador', 'transformador', 'aterramento',
        'instalação elétrica', 'ponto de força', 'caixa de passagem',
    ],
    'Arte / Impressão': [
        'arte', 'artwork', 'layout', 'design', 'criação', 'job',
        'arquivo', 'arquivo gráfico', 'vetor', 'mockup', 'aprovação de arte',
        'impressão', 'gráfica', 'plotagem', 'sublimação', 'bordado',
        'serigrafia', 'silk', 'hot stamping', 'uv',
    ],
    'Locação': [
        'locação', 'locacao', 'aluguel', 'alug', 'locado', 'locar',
        'aluga-se', 'para alugar', 'locatário', 'locador',
    ],
    'Compra': [
        'compra', 'adquirir', 'aquisição', 'aquisicao', 'comprar',
        'solicitar compra', 'pedido de compra', 'cotação', 'cotacao',
        'fornecimento', 'suprimento',
    ],
    'Audiovisual': [
        'audiovisual', 'tv', 'monitor', 'tela', 'projeção', 'projetor',
        'datashow', 'led wall', 'videowall', 'video wall', 'plasma',
        'som', 'áudio', 'audio', 'caixa de som', 'microfone', 'hdmi',
        'streaming', 'câmera', 'camera',
    ],
}

# ── Materiais ─────────────────────────────────────────────────────────────────
MATERIAL_KEYWORDS = {
    'MDF': ['mdf', 'm.d.f'],
    'Madeira': ['madeira', 'pinus', 'eucalipto', 'cedro', 'teca', 'cumaru',
                'compensado', 'eucatex', 'osb', 'plywood', 'compensado naval'],
    'Acrílico': ['acrílico', 'acrilico', 'plexi', 'plexiglas'],
    'Vidro': ['vidro', 'glass', 'espelho', 'laminado de vidro'],
    'Ferro / Aço': ['ferro', 'aço', 'aco', 'inox', 'aço inox', 'aço galvanizado',
                    'chapa de aço', 'perfil de aço'],
    'Alumínio': ['alumínio', 'aluminio', 'perfil de alumínio', 'chapa de alumínio'],
    'PVC': ['pvc', 'tubo pvc', 'cano pvc'],
    'Lona': ['lona', 'banner', 'front light', 'back light', 'lona frontlit',
             'lona backlit', 'lona impressa'],
    'Tecido': ['tecido', 'tecido blackout', 'tecido de malha', 'percal',
               'oxford', 'voil', 'veludo', 'lycra', 'neoprene'],
    'Metalon': ['metalon', 'perfil tubular', 'tubo metalon', 'tubo estrutural'],
    'Concreto': ['concreto', 'cimento', 'argamassa'],
    'Pedra': ['mármore', 'marmore', 'granito', 'quartzito', 'slate', 'ardósia',
              'pedra natural', 'pedra artificial'],
    'Formica / Laminado': ['fórmica', 'formica', 'laminado melamínico',
                           'laminado de alta pressão', 'laminado decorativo'],
    'Gesso': ['gesso', 'drywall', 'sancas', 'moldura de gesso'],
    'Foam / Isopor': ['isopor', 'foam', 'eps', 'xps', 'espuma'],
}

# ── Cores e acabamentos ───────────────────────────────────────────────────────
COLOR_KEYWORDS = [
    'branco', 'preto', 'cinza', 'prata', 'dourado', 'dorado', 'cobre', 'bronze',
    'bege', 'creme', 'off-white', 'off white', 'marrom', 'caramelo', 'tabaco',
    'azul', 'azul royal', 'azul petróleo', 'verde', 'verde militar',
    'vermelho', 'vinho', 'bordô', 'bordo', 'laranja', 'amarelo', 'roxo', 'lilás',
    'rosa', 'salmão', 'coral', 'nude', 'natural', 'madeirado', 'amadeirado',
    'escovado', 'polido', 'fosco', 'brilhante', 'acetinado', 'texturizado',
    'grafite', 'champagne', 'ouro', 'platina', 'antracite',
]

ACABAMENTO_KEYWORDS = [
    'fosco', 'brilhante', 'acetinado', 'semibrilho', 'semi-brilho',
    'polido', 'escovado', 'texturizado', 'laminado', 'laqueado', 'lacado',
    'pintado', 'anodizado', 'galvanizado', 'cromado', 'espelhado',
    'natural', 'bruto', 'envelhecido', 'patinado', 'oxidado',
]

# ── Tipo de produção ──────────────────────────────────────────────────────────
TIPO_PRODUCAO_MAP = {
    'Marcenaria': ['marcenaria', 'marceneiro', 'carpintaria', 'carpinteiro', 'corte e dobra em madeira'],
    'Serralheria': ['serralheria', 'serralheiro', 'soldagem', 'metalúrgica',
                    'corte e dobra', 'dobra em aço', 'metalon'],
    'Impressão Gráfica': ['impressão', 'plotagem', 'gráfica', 'grafica', 'impressa',
                          'imprimir', 'sublimação', 'serigrafia', 'silk'],
    'Adesivação': ['adesivação', 'adesivacao', 'adesivo', 'plotar'],
    'Pintura': ['pintura', 'pintar', 'tinta', 'verniz', 'laca', 'lacar'],
    'Locação': ['locação', 'locacao', 'alugar', 'aluguel', 'locar'],
    'Compra': ['compra', 'comprar', 'adquirir', 'aquisição', 'cotação'],
    'Instalação': ['instalação', 'instalar', 'montagem', 'montar', 'fixação', 'fixar'],
    'Obra Civil': ['obra', 'alvenaria', 'reboco', 'emassamento', 'drywall'],
    'Item Existente': ['existente', 'já existe', 'reutilizar', 'reaproveitar'],
    'A Fabricar': ['fabricar', 'produzir', 'confeccionar', 'executar', 'fazer'],
}

# ── Status de arte ────────────────────────────────────────────────────────────
STATUS_ARTE_KEYWORDS = {
    'Aprovado': ['aprovado', 'aprovada', 'arte aprovada', 'arte ok', 'arte liberada',
                 'confirmado', 'liberado', 'ok', 'fechado', '✓', '✔'],
    'Em Aprovação': ['em aprovação', 'em aprovacao', 'aguardando aprovação',
                     'aguardando aprovacao', 'para aprovação', 'a aprovar',
                     'em análise', 'em analise'],
    'Não Aprovado': ['não aprovado', 'nao aprovado', 'reprovado', 'rejeitado',
                     'refazer', 'revisar', 'revisão'],
    'Aguardando Arte': ['aguardando arte', 'arte pendente', 'a criar', 'a desenvolver',
                        'sem arte', 'arte não recebida', 'arte nao recebida'],
    'A Produzir': ['a produzir', 'produzir', 'em produção', 'em producao',
                   'produzindo', 'na produção'],
    'Pendente': ['pendente', 'pending', 'a definir', 'a confirmar'],
}

STATUS_COMPRA_KEYWORDS = {
    'Comprado': ['comprado', 'adquirido', 'pedido realizado', 'pedido confirmado',
                 'em trânsito', 'entregue'],
    'Locado': ['locado', 'alugado', 'reservado', 'confirmado fornecedor'],
    'A Comprar': ['a comprar', 'comprar', 'solicitar compra', 'pendente compra'],
    'A Locar': ['a locar', 'a alugar', 'buscar locação', 'pesquisar locação'],
    'Cotando': ['cotando', 'cotação', 'em cotação', 'solicitar cotação'],
    'Pendente': ['pendente', 'a confirmar', 'aguardando', 'indefinido'],
}

# ── Palavras-chave de pendência ───────────────────────────────────────────────
PENDENCIA_KEYWORDS = [
    'pendente', 'pendência', 'pendencia', 'a confirmar', 'a definir',
    'aguardando', 'aguardar', 'verificar', 'checar', 'conferir',
    'falta definir', 'não aprovado', 'nao aprovado', 'não definido',
    'sem confirmação', 'sem confirmacao', 'indefinido', 'a resolver',
    'necessário confirmar', 'necessario confirmar', 'a providenciar',
    'solicitar', 'falta definir', 'não definido', 'nao definido',
    'a validar', 'aprovação pendente',
    'arte não recebida', 'arte nao recebida',
]

# ── Nível de atenção ──────────────────────────────────────────────────────────
NIVEL_ALTO_KEYWORDS = [
    'urgente', 'crítico', 'critico', 'bloqueado', 'bloqueante',
    'impede', 'impacto', 'atrasado', 'atraso', 'alerta',
    'erro', 'errado', 'incorreto', 'problema', 'falha',
]
NIVEL_MEDIO_KEYWORDS = [
    'atenção', 'atencao', 'verificar', 'checar', 'revisar', 'revisão',
    'confirmar', 'pendente', 'aguardando', 'dúvida', 'duvida',
]

# ── Mapeamento de colunas de tabela ───────────────────────────────────────────
TABLE_HEADER_SYNONYMS = {
    'item':       ['item', 'nome', 'descrição', 'descricao', 'produto', 'material',
                   'peça', 'peca', 'objeto', 'elemento', 'componente', 'especificação'],
    'quantidade': ['qtd', 'qt', 'quant', 'quantidade', 'qnt', 'qde', 'qtde', 'q.'],
    'unidade':    ['un', 'und', 'unid', 'unidade', 'medida', 'tipo'],
    'medida':     ['medidas', 'dimensões', 'dimensao', 'dimensoes', 'medida',
                   'tamanho', 'l x a', 'l x h', 'formato', 'dimensão'],
    'material':   ['material', 'mat.', 'substrato', 'acabamento', 'composição'],
    'status':     ['status', 'situação', 'situacao', 'andamento', 'estado',
                   'aprovação', 'aprovacao', 'ok', 'situação atual'],
    'observacao': ['obs', 'observação', 'observacao', 'observações', 'observacoes',
                   'nota', 'notas', 'detalhe', 'detalhes', 'especificação'],
    'fornecedor': ['fornecedor', 'empresa', 'responsável', 'responsavel', 'exec.',
                   'executante', 'quem faz', 'fabricante'],
    'categoria':  ['categoria', 'tipo', 'classe', 'setor', 'área'],
    'codigo':     ['código', 'codigo', 'cod', 'ref', 'referência', 'referencia', 'id'],
    'valor':      ['valor', 'preço', 'preco', 'custo', 'vr', 'total', 'unitário'],
    'ambiente':   ['ambiente', 'local', 'espaço', 'espaco', 'área', 'seção', 'secao'],
}

# ── Indicadores de seção / cabeçalho ─────────────────────────────────────────
SECTION_INDICATOR_WORDS = [
    'seção', 'secao', 'área', 'area', 'ambiente', 'bloco', 'módulo', 'modulo',
    'fase', 'etapa', 'parte', 'capítulo', 'capitulo', 'item', 'grupo',
    'cenário', 'cenario', 'espaço', 'espaco', 'salão', 'salao', 'hall',
    'estande', 'stand', 'expositores', 'entrada', 'recepção', 'recepcao',
    'palco', 'stage', 'camarim', 'backstage', 'corredor', 'circulação',
]

# ── Indicadores de fornecedor ─────────────────────────────────────────────────
FORNECEDOR_HINTS = [
    # Words that often precede a supplier name
    'exec.', 'executante:', 'fornecedor:', 'responsável:', 'empresa:',
    'por:', 'feito por', 'produzido por', 'fornecido por',
]
