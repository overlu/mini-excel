<?php
/**
 * This file is part of MiniExcel.
 * @auth lupeng
 */
declare(strict_types=1);

namespace MiniExcel;

use Mini\Support\Collection;
use Mini\Support\Str;
use OpenSpout\Common\Entity\Cell;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Reader\Exception\ReaderNotOpenedException;
use OpenSpout\Reader\ReaderInterface;
use OpenSpout\Reader\SheetInterface;
use OpenSpout\Writer\Common\AbstractOptions;

/**
 * Trait Importable.
 *
 * @property int $start_row
 * @property bool $transpose
 * @property bool $with_header
 */
trait Importable
{
    /**
     * @var int
     */
    private int $sheet_number = 1;

    /**
     * @param AbstractOptions $options
     */
    abstract protected function setOptions(&$options);

    /**
     * @param string $path
     * @param callable|null $callback
     * @return Collection
     * @throws IOException
     * @throws ReaderNotOpenedException
     */
    public function import(string $path, callable $callback = null): Collection
    {
        $reader = $this->reader($path);

        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->sheet_number !== $key) {
                continue;
            }
            $collection = $this->importSheet($sheet, $callback);
        }
        $reader->close();

        return collect($collection ?? []);
    }

    /**
     * @param string $path
     * @param callable|null $callback
     * @return Collection
     * @throws IOException
     * @throws ReaderNotOpenedException
     */
    public function importSheets(string $path, callable $callback = null): Collection
    {
        $reader = $this->reader($path);

        $collections = [];
        foreach ($reader->getSheetIterator() as $key => $sheet) {
            if ($this->with_sheets_names) {
                $collections[$sheet->getName()] = $this->importSheet($sheet, $callback);
            } else {
                $collections[] = $this->importSheet($sheet, $callback);
            }
        }
        $reader->close();

        return new Collection($collections);
    }

    /**
     * @param string $path
     * @return ReaderInterface
     * @throws IOException
     */
    private function reader(string $path): ReaderInterface
    {
        if (Str::endsWith($path, 'csv')) {
            $options = new \OpenSpout\Reader\CSV\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\CSV\Reader($options);
        } elseif (Str::endsWith($path, 'ods')) {
            $options = new \OpenSpout\Reader\ODS\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\ODS\Reader($options);
        } else {
            $options = new \OpenSpout\Reader\XLSX\Options();
            $this->setOptions($options);
            $reader = new \OpenSpout\Reader\XLSX\Reader($options);
        }

        /* @var ReaderInterface $reader */
        $reader->open($path);

        return $reader;
    }

    /**
     * @param array $array
     *
     * @return array
     */
    private function transposeCollection(array $array): array
    {
        $collection = [];

        foreach ($array as $row => $columns) {
            foreach ($columns as $column => $value) {
                data_set(
                    $collection,
                    implode('.', [
                        $column,
                        $row,
                    ]),
                    $value
                );
            }
        }

        return $collection;
    }

    /**
     * @param SheetInterface $sheet
     * @param callable|null $callback
     *
     * @return array
     */
    private function importSheet(SheetInterface $sheet, callable $callback = null): array
    {
        $headers = [];
        $collection = [];
        $count_header = 0;

        foreach ($sheet->getRowIterator() as $k => $rowAsObject) {
            $row = array_map(static function (Cell $cell) {
                return match (true) {
                    $cell instanceof Cell\FormulaCell => $cell->getComputedValue(),
                    default => $cell->getValue(),
                };
            }, $rowAsObject->getCells());

            if ($k >= $this->start_row) {
                if ($this->with_header) {
                    if ($k === $this->start_row) {
                        $headers = $this->toStrings($row);
                        $count_header = count($headers);
                        continue;
                    }
                    if ($count_header > $count_row = count($row)) {
                        $row = array_merge($row, array_fill(0, $count_header - $count_row, null));
                    } elseif ($count_header < count($row)) {
                        $row = array_slice($row, 0, $count_header);
                    }
                }
                if ($callback) {
                    if ($result = $callback(empty($headers) ? $row : array_combine($headers, $row))) {
                        $collection[] = $result;
                    }
                } else {
                    $collection[] = empty($headers) ? $row : array_combine($headers, $row);
                }
            }
        }

        if ($this->transpose) {
            return $this->transposeCollection($collection);
        }

        return $collection;
    }

    /**
     * @param array $values
     *
     * @return array
     */
    private function toStrings(array $values): array
    {
        foreach ($values as &$value) {
            if ($value instanceof \DateTime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value instanceof \DateTimeImmutable) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif ($value) {
                $value = (string)$value;
            }
        }

        return $values;
    }
}
