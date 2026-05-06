"use client";

import { useState, useCallback } from "react";
import * as XLSX from "xlsx";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { Button } from "@/components/ui/button";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { Upload, Users, Building2, Clock, FileSpreadsheet, X } from "lucide-react";

interface EquipeData {
  [equipe: string]: {
    [cliente: string]: number;
  };
}

interface TecnicoData {
  [tecnico: string]: number;
}

// Normaliza texto removendo acentos e convertendo para minúsculas
function normalizar(txt: string): string {
  return txt
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

// Converte valor para segundos - CORRIGIDO para lidar com diferentes formatos do Excel
function converterParaSegundos(valor: unknown): number {
  if (valor === undefined || valor === null || valor === "") return 0;

  // Excel armazena tempo como fração de dia (0.5 = 12 horas)
  if (typeof valor === "number") {
    // Verifica se é um valor de tempo válido (entre 0 e 1 para horas dentro de um dia)
    // ou maior que 1 para durações maiores que 24h
    if (valor >= 0 && valor < 1) {
      // Valor é fração de dia
      return Math.round(valor * 86400);
    } else if (valor >= 1 && valor < 100) {
      // Pode ser horas decimais (ex: 8.5 = 8h30min)
      const horas = Math.floor(valor);
      const minutos = Math.round((valor - horas) * 60);
      return horas * 3600 + minutos * 60;
    } else {
      // Pode ser um serial date do Excel - extrair apenas a parte do tempo
      const fracao = valor % 1;
      return Math.round(fracao * 86400);
    }
  }

  const valorStr = valor.toString().trim();

  // Formato HH:MM:SS ou HH:MM
  if (valorStr.includes(":")) {
    const partes = valorStr.split(":");
    const horas = parseFloat(partes[0]) || 0;
    const minutos = parseFloat(partes[1]) || 0;
    const segundos = parseFloat(partes[2]) || 0;
    return Math.round(horas * 3600 + minutos * 60 + segundos);
  }

  // Tenta converter como número
  const num = parseFloat(valorStr);
  if (!isNaN(num)) {
    if (num >= 0 && num < 1) {
      return Math.round(num * 86400);
    } else if (num >= 1 && num < 100) {
      const horas = Math.floor(num);
      const minutos = Math.round((num - horas) * 60);
      return horas * 3600 + minutos * 60;
    }
  }

  return 0;
}

// Formata segundos para HH:MM:SS
function formatarTempo(seg: number): string {
  seg = Math.round(seg);
  const h = Math.floor(seg / 3600);
  const m = Math.floor((seg % 3600) / 60);
  const s = seg % 60;

  return (
    String(h).padStart(2, "0") +
    ":" +
    String(m).padStart(2, "0") +
    ":" +
    String(s).padStart(2, "0")
  );
}

// Formata segundos para exibição mais amigável
function formatarTempoLegivel(seg: number): string {
  seg = Math.round(seg);
  const h = Math.floor(seg / 3600);
  const m = Math.floor((seg % 3600) / 60);

  if (h === 0) {
    return `${m}min`;
  }
  return m > 0 ? `${h}h ${m}min` : `${h}h`;
}

export function HoursReport() {
  const [porEquipe, setPorEquipe] = useState<EquipeData>({});
  const [porTecnico, setPorTecnico] = useState<TecnicoData>({});
  const [fileName, setFileName] = useState<string>("");
  const [isDragging, setIsDragging] = useState(false);
  const [totalRegistros, setTotalRegistros] = useState(0);

  const processarArquivo = useCallback((file: File) => {
    setFileName(file.name);
    const reader = new FileReader();

    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target?.result as ArrayBuffer);
      const wb = XLSX.read(data, { type: "array" });
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(sheet);

      processarDados(json as Record<string, unknown>[]);
    };

    reader.readAsArrayBuffer(file);
  }, []);

  const processarDados = (linhas: Record<string, unknown>[]) => {
    const novoEquipe: EquipeData = {};
    const novoTecnico: TecnicoData = {};
    let registrosProcessados = 0;

    linhas.forEach((row) => {
      let duracao = 0;
      let cliente = "";
      let equipe = "";

      // Busca o técnico em diferentes variações do nome da coluna
      let tecnico =
        (row["Responsavel"] as string) ||
        (row["Responsável"] as string) ||
        (row["responsavel"] as string) ||
        (row["RESPONSAVEL"] as string) ||
        (row["Técnico"] as string) ||
        (row["Tecnico"] as string) ||
        "";

      Object.keys(row).forEach((col) => {
        const chave = normalizar(col);

        if (chave.includes("duracao") || chave.includes("tempo") || chave.includes("horas")) {
          duracao = converterParaSegundos(row[col]);
        }

        if (chave.includes("cliente") || chave.includes("loja")) {
          cliente = row[col] as string;
        }

        if (chave.includes("equipe")) {
          equipe = row[col] as string;
        }
      });

      if (!cliente || !equipe || !tecnico) return;

      registrosProcessados++;

      // Agrupa por Equipe + Cliente
      if (!novoEquipe[equipe]) novoEquipe[equipe] = {};
      if (!novoEquipe[equipe][cliente]) novoEquipe[equipe][cliente] = 0;
      novoEquipe[equipe][cliente] += duracao;

      // Agrupa por Técnico
      if (!novoTecnico[tecnico]) novoTecnico[tecnico] = 0;
      novoTecnico[tecnico] += duracao;
    });

    setPorEquipe(novoEquipe);
    setPorTecnico(novoTecnico);
    setTotalRegistros(registrosProcessados);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      processarArquivo(file);
    }
  };

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragging(false);
      const file = e.dataTransfer.files?.[0];
      if (file) {
        processarArquivo(file);
      }
    },
    [processarArquivo]
  );

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  }, []);

  const limparDados = () => {
    setPorEquipe({});
    setPorTecnico({});
    setFileName("");
    setTotalRegistros(0);
  };

  // Calcula totais
  const totalEquipe = Object.values(porEquipe).reduce(
    (acc, clientes) =>
      acc + Object.values(clientes).reduce((sum, seg) => sum + seg, 0),
    0
  );

  const totalTecnicos = Object.values(porTecnico).reduce(
    (acc, seg) => acc + seg,
    0
  );

  const temDados = Object.keys(porEquipe).length > 0;

  return (
    <div className="min-h-screen p-6 md:p-8">
      <div className="max-w-7xl mx-auto space-y-8">
        {/* Header */}
        <div className="space-y-2">
          <h1 className="text-3xl md:text-4xl font-bold text-foreground tracking-tight">
            Relatório de Horas
          </h1>
          <p className="text-muted-foreground">
            Analise as horas trabalhadas por equipe, loja e técnico
          </p>
        </div>

        {/* Upload Area */}
        <Card className="border-border bg-card">
          <CardHeader className="pb-4">
            <CardTitle className="flex items-center gap-2 text-lg text-card-foreground">
              <FileSpreadsheet className="h-5 w-5 text-primary" />
              Upload da Planilha
            </CardTitle>
          </CardHeader>
          <CardContent>
            <div
              className={`relative border-2 border-dashed rounded-xl p-8 md:p-12 text-center transition-all duration-200 ${
                isDragging
                  ? "border-primary bg-primary/10"
                  : "border-border hover:border-primary/50 hover:bg-secondary/50"
              }`}
              onDrop={handleDrop}
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
            >
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleFileChange}
                className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
              />
              <div className="space-y-4">
                <div className="mx-auto w-16 h-16 rounded-full bg-secondary flex items-center justify-center">
                  <Upload className="h-8 w-8 text-primary" />
                </div>
                <div>
                  <p className="text-foreground font-medium">
                    Arraste e solte sua planilha aqui
                  </p>
                  <p className="text-sm text-muted-foreground mt-1">
                    ou clique para selecionar (XLSX, XLS, CSV)
                  </p>
                </div>
              </div>
            </div>

            {fileName && (
              <div className="mt-4 flex items-center justify-between p-3 bg-secondary rounded-lg">
                <div className="flex items-center gap-3">
                  <FileSpreadsheet className="h-5 w-5 text-primary" />
                  <div>
                    <p className="text-sm font-medium text-foreground">{fileName}</p>
                    <p className="text-xs text-muted-foreground">
                      {totalRegistros} registros processados
                    </p>
                  </div>
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={limparDados}
                  className="text-muted-foreground hover:text-foreground"
                >
                  <X className="h-4 w-4" />
                </Button>
              </div>
            )}
          </CardContent>
        </Card>

        {temDados && (
          <>
            {/* Stats Cards */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <Card className="border-border bg-card">
                <CardContent className="pt-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-sm text-muted-foreground">Total de Horas</p>
                      <p className="text-2xl font-bold text-foreground mt-1">
                        {formatarTempoLegivel(totalEquipe)}
                      </p>
                    </div>
                    <div className="h-12 w-12 rounded-lg bg-primary/10 flex items-center justify-center">
                      <Clock className="h-6 w-6 text-primary" />
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-border bg-card">
                <CardContent className="pt-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-sm text-muted-foreground">Equipes</p>
                      <p className="text-2xl font-bold text-foreground mt-1">
                        {Object.keys(porEquipe).length}
                      </p>
                    </div>
                    <div className="h-12 w-12 rounded-lg bg-accent/10 flex items-center justify-center">
                      <Building2 className="h-6 w-6 text-accent" />
                    </div>
                  </div>
                </CardContent>
              </Card>

              <Card className="border-border bg-card">
                <CardContent className="pt-6">
                  <div className="flex items-center justify-between">
                    <div>
                      <p className="text-sm text-muted-foreground">Técnicos</p>
                      <p className="text-2xl font-bold text-foreground mt-1">
                        {Object.keys(porTecnico).length}
                      </p>
                    </div>
                    <div className="h-12 w-12 rounded-lg bg-chart-3/10 flex items-center justify-center">
                      <Users className="h-6 w-6 text-chart-3" />
                    </div>
                  </div>
                </CardContent>
              </Card>
            </div>

            {/* Tables Grid */}
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
              {/* Tabela por Equipe e Loja */}
              <Card className="border-border bg-card">
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-card-foreground">
                    <Building2 className="h-5 w-5 text-primary" />
                    Total por Equipe e Loja
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="rounded-lg border border-border overflow-hidden">
                    <Table>
                      <TableHeader>
                        <TableRow className="bg-secondary hover:bg-secondary">
                          <TableHead className="text-foreground font-semibold">Equipe</TableHead>
                          <TableHead className="text-foreground font-semibold">Cliente</TableHead>
                          <TableHead className="text-foreground font-semibold text-right">Horas</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {Object.entries(porEquipe).map(([equipe, clientes]) =>
                          Object.entries(clientes).map(([cliente, seg], idx) => (
                            <TableRow key={`${equipe}-${cliente}`} className="hover:bg-secondary/50">
                              <TableCell className="font-medium text-foreground">
                                {idx === 0 ? equipe : ""}
                              </TableCell>
                              <TableCell className="text-muted-foreground">{cliente}</TableCell>
                              <TableCell className="text-right font-mono text-foreground">
                                {formatarTempo(seg)}
                              </TableCell>
                            </TableRow>
                          ))
                        )}
                        <TableRow className="bg-primary/10 hover:bg-primary/10 font-bold">
                          <TableCell colSpan={2} className="text-foreground">TOTAL</TableCell>
                          <TableCell className="text-right font-mono text-primary">
                            {formatarTempo(totalEquipe)}
                          </TableCell>
                        </TableRow>
                      </TableBody>
                    </Table>
                  </div>
                </CardContent>
              </Card>

              {/* Tabela por Técnico */}
              <Card className="border-border bg-card">
                <CardHeader>
                  <CardTitle className="flex items-center gap-2 text-card-foreground">
                    <Users className="h-5 w-5 text-primary" />
                    Total por Técnico
                  </CardTitle>
                </CardHeader>
                <CardContent>
                  <div className="rounded-lg border border-border overflow-hidden">
                    <Table>
                      <TableHeader>
                        <TableRow className="bg-secondary hover:bg-secondary">
                          <TableHead className="text-foreground font-semibold">Técnico</TableHead>
                          <TableHead className="text-foreground font-semibold text-right">Horas</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {Object.entries(porTecnico)
                          .sort(([, a], [, b]) => b - a)
                          .map(([tecnico, seg]) => (
                            <TableRow key={tecnico} className="hover:bg-secondary/50">
                              <TableCell className="font-medium text-foreground">{tecnico}</TableCell>
                              <TableCell className="text-right font-mono text-foreground">
                                {formatarTempo(seg)}
                              </TableCell>
                            </TableRow>
                          ))}
                        <TableRow className="bg-primary/10 hover:bg-primary/10 font-bold">
                          <TableCell className="text-foreground">TOTAL</TableCell>
                          <TableCell className="text-right font-mono text-primary">
                            {formatarTempo(totalTecnicos)}
                          </TableCell>
                        </TableRow>
                      </TableBody>
                    </Table>
                  </div>
                </CardContent>
              </Card>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
