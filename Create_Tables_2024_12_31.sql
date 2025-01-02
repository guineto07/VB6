/*========================================================================================================================
    AUTOR: Guilherme Neto
    DATA DE CRIAÇÃO: 31/12/2024
    DESCRIÇÃO: 
    Este script cria duas tabelas, uma chamada 'Clientes' e a outra 'Cartao_Transacoes'e faz o relacionamento necessário entre elas. 
	Além disso, ele aprimora a performance ao adicionar índices nas colunas que devem ser mais usadas nas cláusulas WHERE, JOIN ou ORDER BY. 
	O comando DELETE pode ocasionar em fragmentação nos dados, pois embora exclua registros, 
	o espaço onde esses dados estavam alocados nas páginas não é imediatamente otimizado ou reutilizado, levando à fragmentação interna. 
	Para solucionar esse problema, eu optei por incluir o campo 'Status' na tabela 'Cartao_Transacoes', possibilitando a implementação de 'Soft Delete', 
	que evita essa fragmentação e melhora o desempenho geral.
==============================================================================================================================================*/

Use Desenvolvimento;

--	Verifica se a tabela 'Clientes' existe. Se não existir, cria a tabela com os campos necessários.
--	A tabela 'Clientes' possui o campo 'Id_Cliente' como chave primária e 'Numero_Cartao' como único. 

IF OBJECT_ID('Clientes', 'U') IS NULL -- (User Table)
    CREATE TABLE Clientes (
        Id_Cliente INT IDENTITY(1,1) PRIMARY KEY CLUSTERED,   -- Chave primária com incremento automático
        Nome VARCHAR(100) NOT NULL,                   -- Nome do cliente, campo obrigatório
        Numero_Cartao VARCHAR(16) NOT NULL,               -- Número do cartão, campo obrigatório
        CONSTRAINT UQ_Numero_Cartao UNIQUE (Numero_Cartao)  -- Garante que o número do cartão seja único
    );

-- Verifica se a tabela 'Cartao_Transacoes' existe. Se não existir, cria a tabela para armazenar as transações.
IF OBJECT_ID('Cartao_Transacoes', 'U') IS NULL
BEGIN
    CREATE TABLE Cartao_Transacoes (
        Id_Transacao BIGINT IDENTITY(1,1) PRIMARY KEY,      -- Chave primária com incremento automático
        Id_Cliente INT NOT NULL,                         -- Chave estrangeira para a tabela 'Clientes'
        Numero_Cartao VARCHAR(16) NOT NULL,              -- Relacionado ao número do cartão na tabela 'Clientes'
        Valor_Transacao DECIMAL(12, 2) NOT NULL,         -- Valor da transação, campo obrigatório
        Data_Transacao DATETIME NOT NULL,                -- Data da transação, campo obrigatório
        Descricao VARCHAR(100) NULL,                     -- Descrição opcional da transação
        Status CHAR(1) DEFAULT 'A' -- A = Ativo, I = Inativo, padrão 'A' para ativo
    );
END;

GO

-- Criação de chave estrangeira entre 'Cartao_Transacoes' e 'Clientes' para o campo 'Id_Cliente'
-- Isso garante que as transações só poderão ser associadas a um cliente válido.
ALTER TABLE Cartao_Transacoes
ADD CONSTRAINT FK_Cartao_Transacoes_Clientes
FOREIGN KEY (Id_Cliente) REFERENCES Clientes(Id_Cliente)

-- Criação de chave estrangeira entre 'Cartao_Transacoes' e 'Clientes' para o campo 'Numero_Cartao'
-- Isso garante que as transações só poderão ser associadas a um número de cartão válido.
ALTER TABLE Cartao_Transacoes
ADD CONSTRAINT FK_Cartao_Transacoes_Numero_Cartao
FOREIGN KEY (Numero_Cartao) REFERENCES Clientes(Numero_Cartao)
GO

-- Criação de índices para melhorar a performance das consultas na tabela 'Cartao_Transacoes'.
-- O índice 'IDX_Id_Cliente' melhora a performance nas consultas filtrando por 'Id_Cliente'.
CREATE INDEX IDX_Id_Cliente ON Cartao_Transacoes(Id_Cliente);

-- O índice 'IDX_Numero_Cartao' melhora a performance nas consultas filtrando por 'Numero_Cartao'.
CREATE INDEX IDX_Numero_Cartao ON Cartao_Transacoes(Numero_Cartao);

-- O índice 'IDX_Data_Transacao' melhora a performance nas consultas por 'Data_Transacao'.
CREATE INDEX IDX_Data_Transacao ON Cartao_Transacoes(Data_Transacao);

--FIM
GO
--=================================================================================================================================
-- Insere dados para testes
INSERT INTO Clientes VALUES ('GUILHERME ALVES NETO', '1111111111111111'); -- Autoincremento, id_cliente = 1
INSERT INTO Clientes VALUES ('GABRIELA RODRIGUES PEREIRA NETO', '2222222222222222'); -- id-cliente = 2

select * from Clientes
select * from Cartao_Transacoes

