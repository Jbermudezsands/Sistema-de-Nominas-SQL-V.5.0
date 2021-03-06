
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[IndiceTicket](
	[NumeroFG] [nvarchar](50) NOT NULL,
	[NumeroMesa] [float] NOT NULL,
	[Color] [nvarchar](50) NOT NULL,
	[FechaInicial] [datetime] NULL,
	[NumeroBultos] [int] NULL,
	[Serie] [nvarchar](50) NULL,
 CONSTRAINT [PK_IndiceTicket] PRIMARY KEY CLUSTERED 
(
	[NumeroFG] ASC,
	[NumeroMesa] ASC,
	[Color] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]
