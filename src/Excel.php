<?php
/**
 * This file is part of MiniExcel.
 * @auth lupeng
 */
declare(strict_types=1);

namespace MiniExcel;

use Generator;
use Mini\Support\Collection;
use OpenSpout\Reader\CSV\Options as CsvReaderOptions;
use OpenSpout\Reader\XLSX\Options;
use OpenSpout\Writer\CSV\Options as CsvWriterOptions;
use OpenSpout\Reader\ODS\Options as OdsReaderOptions;
use OpenSpout\Writer\ODS\Options as OdsWriterOptions;

/**
 * Class FastExcel.
 */
class Excel
{
    use Importable;
    use Exportable;

    /**
     * @var Collection|Generator|array|null
     */
    protected Collection|Generator|array|null $data = null;

    /**
     * @var bool
     */
    private bool $with_header = true;

    /**
     * @var bool
     */
    private bool $with_sheets_names = false;

    /**
     * @var int
     */
    private int $start_row = 1;

    /**
     * @var bool
     */
    private bool $transpose = false;

    /**
     * @var array
     */
    private array $csv_configuration = [
        'delimiter' => ',',
        'enclosure' => '"',
        'encoding' => 'UTF-8',
        'bom' => true,
    ];

    /**
     * @var callable
     */
    protected $options_configurator = null;

    /**
     * FastExcel constructor.
     * @param array|Generator|Collection|null $data
     */
    public function __construct(array|Generator|Collection $data = null)
    {
        $this->data = $data;
    }

    /**
     * Manually set data apart from the constructor.
     * @param array|Generator|Collection $data
     * @return $this
     */
    public function data(Collection|array|Generator $data): self
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @param int $sheet_number
     * @return $this
     */
    public function sheet(int $sheet_number): self
    {
        $this->sheet_number = $sheet_number;

        return $this;
    }

    /**
     * @return $this
     */
    public function withoutHeaders(): self
    {
        $this->with_header = false;

        return $this;
    }

    /**
     * @return $this
     */
    public function withSheetsNames(): self
    {
        $this->with_sheets_names = true;

        return $this;
    }

    /**
     * @return $this
     */
    public function startRow(int $row): self
    {
        $this->start_row = $row;

        return $this;
    }

    /**
     * @return $this
     */
    public function transpose(): self
    {
        $this->transpose = true;

        return $this;
    }

    /**
     * @param string $delimiter
     * @param string $enclosure
     * @param string $encoding
     * @param bool $bom
     * @return $this
     */
    public function configureCsv(string $delimiter = ',', string $enclosure = '"', string $encoding = 'UTF-8', bool $bom = false): self
    {
        $this->csv_configuration = compact('delimiter', 'enclosure', 'encoding', 'bom');

        return $this;
    }

    /**
     * Configure the underlying Spout Reader using a callback.
     * @param callable|null $callback
     * @return $this
     * @see configureOptionsUsing
     */
    public function configureReaderUsing(?callable $callback = null): self
    {
        return $this;
    }

    /**
     * Configure the underlying Spout Reader using a callback.
     * @param callable|null $callback
     * @return $this
     * @see configureOptionsUsing
     */
    public function configureWriterUsing(?callable $callback = null): self
    {
        return $this;
    }

    /**
     * Configure the underlying Spout Reader options using a callback.
     * @param callable|null $callback
     * @return $this
     */
    public function configureOptionsUsing(?callable $callback = null): self
    {
        $this->options_configurator = $callback;

        return $this;
    }

    /**
     * @param CsvReaderOptions|CsvWriterOptions|OdsReaderOptions|OdsWriterOptions|Options $options
     */
    protected function setOptions(&$options): void
    {
        if ($options instanceof CsvReaderOptions || $options instanceof CsvWriterOptions) {
            $options->FIELD_DELIMITER = $this->csv_configuration['delimiter'];
            $options->FIELD_ENCLOSURE = $this->csv_configuration['enclosure'];
            if ($options instanceof CsvReaderOptions) {
                $options->ENCODING = $this->csv_configuration['encoding'];
            }
            if ($options instanceof CsvWriterOptions) {
                $options->SHOULD_ADD_BOM = $this->csv_configuration['bom'];
            }
        }

        if (is_callable($this->options_configurator)) {
            call_user_func(
                $this->options_configurator,
                $options
            );
        }
    }
}
