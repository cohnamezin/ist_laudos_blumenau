-- =====================================================================
-- MODELO FÍSICO - PostgreSQL 13+ (gen_random_uuid() nativo)
-- Baseado no DICIONÁRIO DE DADOS | MODELO SQL (12 tabelas, 67 campos, 12 FKs)
-- Projeto: Automação de Laudos e Conformidade Legal Ambiental (IST Blumenau)
--
-- Seções:
--   1. Criação das tabelas (fiel ao dicionário de dados)
--   2. Índices nas chaves estrangeiras (recomendação do dicionário, seção 5)
--   3. Comentários de tabela/coluna que ligam o modelo às regras RN-01 a RN-12
-- Em PostgreSQL 12 ou inferior, habilite antes:  CREATE EXTENSION pgcrypto;
-- =====================================================================

-- ---------------------------------------------------------------------
-- LIMPEZA (opcional)
-- ---------------------------------------------------------------------
DROP TABLE IF EXISTS ensaio_resultado          CASCADE;
DROP TABLE IF EXISTS item_proposta_comercial   CASCADE;
DROP TABLE IF EXISTS item_legislacao           CASCADE;
DROP TABLE IF EXISTS amostra                   CASCADE;
DROP TABLE IF EXISTS proposta_comercial        CASCADE;
DROP TABLE IF EXISTS classe_legislacao         CASCADE;
DROP TABLE IF EXISTS legislacao                CASCADE;
DROP TABLE IF EXISTS unidade_medida            CASCADE;
DROP TABLE IF EXISTS ensaio                    CASCADE;
DROP TABLE IF EXISTS usuario                   CASCADE;
DROP TABLE IF EXISTS laboratorio               CASCADE;
DROP TABLE IF EXISTS cliente                   CASCADE;

-- =====================================================================
-- 1. TABELAS
-- =====================================================================

CREATE TABLE laboratorio (
    id_laboratorio    UUID         NOT NULL DEFAULT gen_random_uuid(),
    sigla_celula      VARCHAR(20)  NOT NULL,
    nome_laboratorio  VARCHAR(100) NOT NULL,
    CONSTRAINT pk_laboratorio PRIMARY KEY (id_laboratorio),
    CONSTRAINT uq_laboratorio_sigla UNIQUE (sigla_celula)
);

CREATE TABLE usuario (
    id_usuario         UUID         NOT NULL DEFAULT gen_random_uuid(),
    nome_usuario       VARCHAR(150) NOT NULL,
    login_sistema      VARCHAR(100) NOT NULL,
    perfil_acesso      VARCHAR(50)  NOT NULL,
    fk_id_laboratorio  UUID         NOT NULL,
    CONSTRAINT pk_usuario PRIMARY KEY (id_usuario),
    CONSTRAINT uq_usuario_login UNIQUE (login_sistema),
    CONSTRAINT fk_usuario_laboratorio
        FOREIGN KEY (fk_id_laboratorio) REFERENCES laboratorio (id_laboratorio)
        ON DELETE RESTRICT
);

CREATE TABLE cliente (
    id_cliente    UUID         NOT NULL DEFAULT gen_random_uuid(),
    razao_social  VARCHAR(150) NOT NULL,
    cnpj_cpf      VARCHAR(18)  NOT NULL,
    email         VARCHAR(150),
    telefone      VARCHAR(20),
    CONSTRAINT pk_cliente PRIMARY KEY (id_cliente),
    CONSTRAINT uq_cliente_documento UNIQUE (cnpj_cpf)
);

CREATE TABLE proposta_comercial (
    id_proposta_comercial   UUID        NOT NULL DEFAULT gen_random_uuid(),
    numero_proposta         VARCHAR(30) NOT NULL,
    qtd_amostras_previstas  INTEGER     NOT NULL DEFAULT 0,
    qtd_ensaios_previstos   INTEGER     NOT NULL DEFAULT 0,
    frequencia_atendimento  VARCHAR(50),
    data_inicio             DATE,
    data_fim                DATE,
    fk_id_cliente           UUID        NOT NULL,
    CONSTRAINT pk_proposta_comercial PRIMARY KEY (id_proposta_comercial),
    CONSTRAINT uq_proposta_numero UNIQUE (numero_proposta),
    CONSTRAINT ck_proposta_qtd_amostras CHECK (qtd_amostras_previstas >= 0),
    CONSTRAINT ck_proposta_qtd_ensaios  CHECK (qtd_ensaios_previstos  >= 0),
    CONSTRAINT ck_proposta_periodo
        CHECK (data_inicio IS NULL OR data_fim IS NULL OR data_fim >= data_inicio),
    CONSTRAINT fk_proposta_cliente
        FOREIGN KEY (fk_id_cliente) REFERENCES cliente (id_cliente)
        ON DELETE RESTRICT
);

CREATE TABLE amostra (
    id_amostra                UUID         NOT NULL DEFAULT gen_random_uuid(),
    codigo_amostra            VARCHAR(50)  NOT NULL,
    data_coleta               TIMESTAMPTZ,
    data_recebimento          TIMESTAMPTZ,
    origem_amostra            VARCHAR(255),
    descricao_do_ponto        VARCHAR(255),
    data_conclusao_amostra    TIMESTAMPTZ,
    data_liberacao            TIMESTAMPTZ,
    fk_id_proposta_comercial  UUID         NOT NULL,
    fk_id_usuario             UUID         NOT NULL,
    CONSTRAINT pk_amostra PRIMARY KEY (id_amostra),
    CONSTRAINT uq_amostra_codigo UNIQUE (codigo_amostra),
    CONSTRAINT ck_amostra_recebimento
        CHECK (data_coleta IS NULL OR data_recebimento IS NULL
               OR data_recebimento >= data_coleta),
    CONSTRAINT ck_amostra_conclusao
        CHECK (data_recebimento IS NULL OR data_conclusao_amostra IS NULL
               OR data_conclusao_amostra >= data_recebimento),
    CONSTRAINT ck_amostra_liberacao
        CHECK (data_conclusao_amostra IS NULL OR data_liberacao IS NULL
               OR data_liberacao >= data_conclusao_amostra),
    CONSTRAINT fk_amostra_proposta
        FOREIGN KEY (fk_id_proposta_comercial) REFERENCES proposta_comercial (id_proposta_comercial)
        ON DELETE RESTRICT,
    CONSTRAINT fk_amostra_usuario
        FOREIGN KEY (fk_id_usuario) REFERENCES usuario (id_usuario)
        ON DELETE RESTRICT
);

CREATE TABLE ensaio (
    id_ensaio      UUID         NOT NULL DEFAULT gen_random_uuid(),
    nome_ensaio    VARCHAR(100) NOT NULL,
    descricao      VARCHAR(255),
    status         BOOLEAN      NOT NULL DEFAULT TRUE,
    codigo_ensaio  INTEGER      NOT NULL,
    CONSTRAINT pk_ensaio PRIMARY KEY (id_ensaio),
    CONSTRAINT uq_ensaio_codigo UNIQUE (codigo_ensaio)
);

CREATE TABLE legislacao (
    id_legislacao        UUID         NOT NULL DEFAULT gen_random_uuid(),
    nome_legislacao      VARCHAR(100) NOT NULL,
    status               BOOLEAN      NOT NULL DEFAULT TRUE,
    data_inicio_vigencia TIMESTAMPTZ,
    data_fim_vigencia    TIMESTAMPTZ,
    CONSTRAINT pk_legislacao PRIMARY KEY (id_legislacao),
    CONSTRAINT ck_legislacao_vigencia
        CHECK (data_inicio_vigencia IS NULL OR data_fim_vigencia IS NULL
               OR data_fim_vigencia >= data_inicio_vigencia)
);

CREATE TABLE classe_legislacao (
    id_classe_legislacao  UUID         NOT NULL DEFAULT gen_random_uuid(),
    nome_classe           VARCHAR(100) NOT NULL,
    status                BOOLEAN      NOT NULL DEFAULT TRUE,
    fk_id_legislacao      UUID         NOT NULL,
    CONSTRAINT pk_classe_legislacao PRIMARY KEY (id_classe_legislacao),
    CONSTRAINT fk_classe_legislacao
        FOREIGN KEY (fk_id_legislacao) REFERENCES legislacao (id_legislacao)
        ON DELETE RESTRICT
);

CREATE TABLE unidade_medida (
    id_unidade_medida  UUID        NOT NULL DEFAULT gen_random_uuid(),
    nome               VARCHAR(50) NOT NULL,
    sigla              VARCHAR(20) NOT NULL,
    status             BOOLEAN     NOT NULL DEFAULT TRUE,
    CONSTRAINT pk_unidade_medida PRIMARY KEY (id_unidade_medida),
    CONSTRAINT uq_unidade_sigla UNIQUE (sigla)
);

-- Sem UNIQUE em (classe, ensaio) de propósito: o legado do UltraLIMS possui
-- duplicidades com VMPs divergentes (RN-04) que precisam ser importadas.
CREATE TABLE item_legislacao (
    id_item_legislacao    UUID          NOT NULL DEFAULT gen_random_uuid(),
    valor_min_permitido   NUMERIC(18,6),
    valor_max_permitido   NUMERIC(18,6),
    fk_id_classe_legislacao  UUID       NOT NULL,
    fk_id_ensaio             UUID       NOT NULL,
    fk_id_unidade_medida     UUID       NOT NULL,
    CONSTRAINT pk_item_legislacao PRIMARY KEY (id_item_legislacao),
    CONSTRAINT ck_item_legislacao_limites
        CHECK (valor_min_permitido IS NULL OR valor_max_permitido IS NULL
               OR valor_max_permitido >= valor_min_permitido),
    CONSTRAINT fk_item_legislacao_classe
        FOREIGN KEY (fk_id_classe_legislacao) REFERENCES classe_legislacao (id_classe_legislacao)
        ON DELETE RESTRICT,
    CONSTRAINT fk_item_legislacao_ensaio
        FOREIGN KEY (fk_id_ensaio) REFERENCES ensaio (id_ensaio)
        ON DELETE RESTRICT,
    CONSTRAINT fk_item_legislacao_unidade
        FOREIGN KEY (fk_id_unidade_medida) REFERENCES unidade_medida (id_unidade_medida)
        ON DELETE RESTRICT
);

CREATE TABLE item_proposta_comercial (
    id_item_proposta_comercial  UUID NOT NULL DEFAULT gen_random_uuid(),
    fk_id_item_legislacao       UUID NOT NULL,
    fk_id_proposta_comercial    UUID NOT NULL,
    CONSTRAINT pk_item_proposta_comercial PRIMARY KEY (id_item_proposta_comercial),
    CONSTRAINT uq_proposta_item_legislacao
        UNIQUE (fk_id_proposta_comercial, fk_id_item_legislacao),
    CONSTRAINT fk_item_proposta_legislacao
        FOREIGN KEY (fk_id_item_legislacao) REFERENCES item_legislacao (id_item_legislacao)
        ON DELETE RESTRICT,
    CONSTRAINT fk_item_proposta_proposta
        FOREIGN KEY (fk_id_proposta_comercial) REFERENCES proposta_comercial (id_proposta_comercial)
        ON DELETE RESTRICT
);

CREATE TABLE ensaio_resultado (
    id_ensaio_resultado     UUID          NOT NULL DEFAULT gen_random_uuid(),
    valor_numerico          NUMERIC(18,6),
    valor_booleano          BOOLEAN,
    metodologia_de_analise  VARCHAR(255),
    data_inicio_analise     TIMESTAMPTZ,
    data_conclusao_analise  TIMESTAMPTZ,
    status_conformidade     BOOLEAN,           -- TRUE conforme | FALSE não conforme | NULL sem avaliação
    fk_id_amostra           UUID          NOT NULL,
    fk_id_item_legislacao   UUID          NOT NULL,
    CONSTRAINT pk_ensaio_resultado PRIMARY KEY (id_ensaio_resultado),
    CONSTRAINT ck_resultado_tipo_valor
        CHECK (valor_numerico IS NULL OR valor_booleano IS NULL),
    CONSTRAINT ck_resultado_periodo
        CHECK (data_inicio_analise IS NULL OR data_conclusao_analise IS NULL
               OR data_conclusao_analise >= data_inicio_analise),
    CONSTRAINT fk_resultado_amostra
        FOREIGN KEY (fk_id_amostra) REFERENCES amostra (id_amostra)
        ON DELETE RESTRICT,
    CONSTRAINT fk_resultado_item_legislacao
        FOREIGN KEY (fk_id_item_legislacao) REFERENCES item_legislacao (id_item_legislacao)
        ON DELETE RESTRICT
);

-- =====================================================================
-- 2. ÍNDICES NAS CHAVES ESTRANGEIRAS (PostgreSQL não os cria sozinho)
-- =====================================================================
CREATE INDEX idx_usuario_laboratorio       ON usuario (fk_id_laboratorio);
CREATE INDEX idx_proposta_cliente          ON proposta_comercial (fk_id_cliente);
CREATE INDEX idx_amostra_proposta          ON amostra (fk_id_proposta_comercial);
CREATE INDEX idx_amostra_usuario           ON amostra (fk_id_usuario);
CREATE INDEX idx_amostra_data_coleta       ON amostra (data_coleta);   -- RN-08 / RN-10
CREATE INDEX idx_classe_legislacao         ON classe_legislacao (fk_id_legislacao);
CREATE INDEX idx_item_legislacao_classe    ON item_legislacao (fk_id_classe_legislacao);
CREATE INDEX idx_item_legislacao_ensaio    ON item_legislacao (fk_id_ensaio);
CREATE INDEX idx_item_legislacao_unidade   ON item_legislacao (fk_id_unidade_medida);
CREATE INDEX idx_item_proposta_item        ON item_proposta_comercial (fk_id_item_legislacao);
CREATE INDEX idx_resultado_amostra         ON ensaio_resultado (fk_id_amostra);
CREATE INDEX idx_resultado_item_legislacao ON ensaio_resultado (fk_id_item_legislacao);

-- =====================================================================
-- 3. COMENTÁRIOS (rastreabilidade com as regras de negócio)
-- =====================================================================
COMMENT ON TABLE  item_legislacao IS 'Tríade Ensaio + Classe/Legislação + Unidade com VMP mín/máx (RN-01). Duplicidades legadas são permitidas e tratadas pela RN-04 (limite mais restritivo).';
COMMENT ON TABLE  item_proposta_comercial IS 'Define quais itens de legislação valem para cada proposta (RN-02 e RN-06: até 3 legislações simultâneas).';
COMMENT ON COLUMN legislacao.data_inicio_vigencia IS 'RN-08: data_inicio <= data_coleta <= data_fim.';
COMMENT ON COLUMN ensaio_resultado.status_conformidade IS 'RN-01/RN-07: TRUE conforme, FALSE não conforme, NULL ainda sem avaliação.';
COMMENT ON COLUMN ensaio_resultado.valor_numerico IS 'Resultado numérico já normalizado (RN-11 conversão de unidade e RN-12 operadores < e >).';
COMMENT ON COLUMN unidade_medida.sigla IS 'Ex.: mg/L, µg/L, UFC/100mL, NMP/100mL, Intensidade (RN-11).';
