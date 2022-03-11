
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 03/11/2022 12:04:51
-- Generated from EDMX file: C:\Users\abhis\source\repos\FIT5122\FIT5122\Models\ProductPackaging.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [Database1];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------


-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------


-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'ProductPackagings'
CREATE TABLE [dbo].[ProductPackagings] (
    [Id] int IDENTITY(1,1) NOT NULL,
    [Productname] nvarchar(max)  NULL,
    [Packaging] nvarchar(max)  NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'ProductPackagings'
ALTER TABLE [dbo].[ProductPackagings]
ADD CONSTRAINT [PK_ProductPackagings]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------