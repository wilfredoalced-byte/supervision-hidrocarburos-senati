CREATE TABLE IF NOT EXISTS registros (
    id INTEGER PRIMARY KEY,
    mes_presenta TEXT,
    oficina_regional TEXT,
    accion TEXT,
    accion_especifica TEXT,
    sub_accion_especifica TEXT,
    tipo_agente TEXT,
    actividad TEXT,
    provincia TEXT,
    distrito TEXT,
    tipo_transporte TEXT,
    numero_supervisores INTEGER CHECK (numero_supervisores > 0),
    empresa_supervisora TEXT,
    calidad_entregable TEXT,
    expediente_siged TEXT,
    carta_linea TEXT,
    observaciones TEXT
);

CREATE INDEX IF NOT EXISTS idx_registros_expediente_siged
    ON registros (expediente_siged);
CREATE INDEX IF NOT EXISTS idx_registros_accion
    ON registros (accion);
CREATE INDEX IF NOT EXISTS idx_registros_empresa_supervisora
    ON registros (empresa_supervisora);
CREATE INDEX IF NOT EXISTS idx_registros_provincia
    ON registros (provincia);
CREATE INDEX IF NOT EXISTS idx_registros_mes_presenta
    ON registros (mes_presenta);
