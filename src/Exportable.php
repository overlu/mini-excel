<?php
/**
 * This file is part of MiniExcel.
 * @auth lupeng
 */
declare(strict_types=1);

namespace MiniExcel;

use Generator;
use Mini\Support\Collection;
use Mini\Support\Str;
use InvalidArgumentException;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Common\Entity\Style\Style;
use OpenSpout\Common\Exception\IOException;
use OpenSpout\Writer\Common\AbstractOptions;
use OpenSpout\Writer\Exception\InvalidSheetNameException;
use OpenSpout\Writer\Exception\WriterNotOpenedException;
use OpenSpout\Writer\WriterInterface;
use Symfony\Component\HttpFoundation\StreamedResponse;

/**
 * Trait Exportable.
 *
 * @property bool $transpose
 * @property bool $with_header
 * @property Collection $data
 */
trait Exportable
{
    private ?Style $header_style = null;
    private ?Style $rows_style = null;

    /**
     * @param AbstractOptions $options
     */
    abstract protected function setOptions(&$options);

    /**
     * @param string $path
     * @param callable|null $callback
     * @return string
     * @throws IOException
     * @throws InvalidSheetNameException
     * @throws WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    public function export(string $path, callable $callback = null): string
    {
        $this->exportOrDownload($path, 'openToFile', $callback);

        return realpath($path) ?: $path;
    }

    /**
     * @param string $path
     * @param callable|null $callback
     * @return StreamedResponse|string
     * @throws IOException
     * @throws InvalidSheetNameException
     * @throws WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    public function download(string $path, callable $callback = null): StreamedResponse|string
    {
        if (method_exists(response(), 'streamDownload')) {
            return response()->streamDownload(function () use ($path, $callback) {
                $this->exportOrDownload($path, 'openToBrowser', $callback);
            }, $path);
        }
        $this->exportOrDownload($path, 'openToBrowser', $callback);

        return '';
    }

    /**
     * @param string $path
     * @param string $function
     * @param callable|null $callback
     * @return void
     * @throws IOException
     * @throws InvalidSheetNameException
     * @throws WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    private function exportOrDownload(string $path, string $function, callable $callback = null): void
    {
        if (Str::endsWith($path, 'csv')) {
            $options = new \OpenSpout\Writer\CSV\Options();
            $writer = new \OpenSpout\Writer\CSV\Writer($options);
        } elseif (Str::endsWith($path, 'ods')) {
            $options = new \OpenSpout\Writer\ODS\Options();
            $writer = new \OpenSpout\Writer\ODS\Writer($options);
        } else {
            $options = new \OpenSpout\Writer\XLSX\Options();
            $writer = new \OpenSpout\Writer\XLSX\Writer($options);
        }

        $this->setOptions($options);
        /* @var WriterInterface $writer */
        $writer->$function($path);

        $has_sheets = ($writer instanceof \OpenSpout\Writer\XLSX\Writer || $writer instanceof \OpenSpout\Writer\ODS\Writer);

        // It can export one sheet (Collection) or N sheets (SheetCollection)
        $data = $this->transpose ? $this->transposeData() : ($this->data instanceof SheetCollection ? $this->data : collect([$this->data]));

        foreach ($data as $key => $collection) {
            if ($collection instanceof Collection) {
                $this->writeRowsFromCollection($writer, $collection, $callback);
            } elseif ($collection instanceof Generator) {
                $this->writeRowsFromGenerator($writer, $collection, $callback);
            } elseif (is_array($collection)) {
                $this->writeRowsFromArray($writer, $collection, $callback);
            } else {
                throw new InvalidArgumentException('Unsupported type for $data');
            }
            if (is_string($key)) {
                $writer->getCurrentSheet()->setName($key);
            }
            if ($has_sheets && $data->keys()->last() !== $key) {
                $writer->addNewSheetAndMakeItCurrent();
            }
        }
        $writer->close();
    }

    /**
     * Transpose data from rows to columns.
     */
    private function transposeData(): Collection
    {
        $data = $this->data instanceof Collection ? $this->data : collect([$this->data]);
        $transposedData = [];

        foreach ($data as $key => $collection) {
            foreach ($collection as $row => $columns) {
                foreach ($columns as $column => $value) {
                    data_set(
                        $transposedData,
                        implode('.', [
                            $key,
                            $column,
                            $row,
                        ]),
                        $value
                    );
                }
            }
        }

        return new Collection($transposedData);
    }

    /**
     * @param WriterInterface $writer
     * @param Collection $collection
     * @param callable|null $callback
     * @return void
     * @throws IOException
     * @throws WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    private function writeRowsFromCollection(WriterInterface $writer, Collection $collection, ?callable $callback = null): void
    {
        // Apply callback
        if ($callback) {
            $collection->transform(function ($value) use ($callback) {
                return $callback($value);
            });
        }
        // Prepare collection (i.e remove non-string)
        $this->prepareCollection($collection);
        // Add header row.
        if ($this->with_header) {
            $this->writeHeader($writer, $collection->first());
        }

        // createRowFromArray works only with arrays
        if (!is_array($collection->first())) {
            $collection = $collection->map(function ($value) {
                return $value->toArray();
            });
        }

        // is_array($first_row) ? $first_row : $first_row->toArray())
        $all_rows = $collection->map(function ($value) {
            return Row::fromValues($value);
        })->toArray();
        if ($this->rows_style) {
            $this->addRowsWithStyle($writer, $all_rows, $this->rows_style);
        } else {
            $writer->addRows($all_rows);
        }
    }

    /**
     * @param WriterInterface $writer
     * @param $all_rows
     * @param $rows_style
     * @return void
     * @throws IOException
     * @throws WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    private function addRowsWithStyle(WriterInterface $writer, $all_rows, $rows_style): void
    {
        $styled_rows = [];
        // Style rows one by one
        foreach ($all_rows as $row) {
            $styled_rows[] = Row::fromValues($row->toArray(), $rows_style);
        }
        $writer->addRows($styled_rows);
    }

    /**
     * @param WriterInterface $writer
     * @param Generator $generator
     * @param callable|null $callback
     * @return void
     * @throws IOException
     * @throws WriterNotOpenedException
     */
    private function writeRowsFromGenerator(WriterInterface $writer, Generator $generator, ?callable $callback = null): void
    {
        foreach ($generator as $key => $item) {
            // Apply callback
            if ($callback) {
                $item = $callback($item);
            }

            // Prepare row (i.e remove non-string)
            $item = $this->transformRow($item);

            // Add header row.
            if ($this->with_header && $key === 0) {
                $this->writeHeader($writer, $item);
            }
            // Write rows (one by one).
            $writer->addRow(Row::fromValues($item->toArray(), $this->rows_style));
        }
    }

    /**
     * @param WriterInterface $writer
     * @param array $array
     * @param callable|null $callback
     * @return void
     * @throws IOException
     * @throws WriterNotOpenedException
     * @throws \OpenSpout\Common\Exception\InvalidArgumentException
     */
    private function writeRowsFromArray(WriterInterface $writer, array $array, ?callable $callback = null): void
    {
        $collection = collect($array);

        if (is_object($collection->first()) || is_array($collection->first())) {
            // provided $array was valid and could be converted to a collection
            $this->writeRowsFromCollection($writer, $collection, $callback);
        }
    }

    /**
     * @param WriterInterface $writer
     * @param $first_row
     * @return void
     * @throws IOException
     * @throws WriterNotOpenedException
     */
    private function writeHeader(WriterInterface $writer, $first_row): void
    {
        if ($first_row === null) {
            return;
        }

        $keys = array_keys(is_array($first_row) ? $first_row : $first_row->toArray());
        $writer->addRow(Row::fromValues($keys, $this->header_style));
    }

    /**
     * @param Collection $collection
     * @return void
     */
    protected function prepareCollection(Collection $collection): void
    {
        $need_conversion = false;
        $first_row = $collection->first();

        if (!$first_row) {
            return;
        }

        foreach ($first_row as $item) {
            if (!is_string($item)) {
                $need_conversion = true;
            }
        }
        if ($need_conversion) {
            $this->transform($collection);
        }
    }

    /**
     * @param Collection $collection
     * @return void
     */
    private function transform(Collection $collection): void
    {
        $collection->transform(function ($data) {
            return $this->transformRow($data);
        });
    }

    /**
     * @param $data
     * @return Collection
     */
    private function transformRow($data): Collection
    {
        return collect($data)->map(function ($value) {
            return is_null($value) ? (string)$value : $value;
        })->filter(function ($value) {
            return is_string($value) || is_int($value) || is_float($value);
        });
    }

    /**
     * @param Style $style
     * @return Excel|Exportable
     */
    public function headerStyle(Style $style): self
    {
        $this->header_style = $style;

        return $this;
    }

    /**
     * @param Style $style
     * @return Excel|Exportable
     */
    public function rowsStyle(Style $style): self
    {
        $this->rows_style = $style;

        return $this;
    }
}
