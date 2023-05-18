<?php

declare(strict_types=1);

namespace SykesCottages\NewZealandOwnerTaxService\Writer;

use DateInterval;
use DatePeriod;
use DateTimeImmutable;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use function array_column;
use function array_filter;
use function array_key_exists;
use function array_shift;
use function array_unique;
use function count;
use function date;
use function number_format;
use function reset;
use function sprintf;
use function str_contains;
use function strlen;
use function strtolower;
use function substr;

use const ARRAY_FILTER_USE_KEY;

class XlsWriter implements Writer
{
    private const DATE_START_ROW_NUMBER       = 8;
    private const DAYS_IN_MONTH               = 31;
    private const TOTAL_ROW_NUMBER            = 33;
    private const MONTH_BLOCK_HEAD_ROWS       = 9;
    private const MONTH_BLOCK_ROWS            = 42;
    private const MONTH_START_COLUMN_POSITION = 2;
    private const COLUMN_TO_OFFEST            = 2;

    private const SUMMARY_START_ROW_NUMBER = 46;

    private const SUMMARY_CATEGORY_NAME_COLUMN = 2;
    private const SUMMARY_TOTALS_COLUMN        = 4;
    private const SUMMARY_DAYS_COLUMN          = 4;
    private const SUMMARY_VALUE_COLUMN         = 5;

    private const SUMMARY_ROW_COUNT_PER_CATEGORY = 2;
    private const CATEGORY_OFFSET                = 2;

    /** @var array|string[] */
    private array $monthsSubheadLabels;
    private float $allValueTotal;
    private int $allDaysTotal;

    /** @var array|string[][]  */
    private array $categories;

    /** @var array|bool[][]  */
    private array $headingStyles;
    /** @var string[]   */
    private array $totalStyles;
    /** @var string[]  */
    private array $numberCurrencyStyle;
    /** @var string[]  */
    private array $borderStyle;

    private Worksheet $activeWorksheet;
    private Spreadsheet $spreadsheet;

    private const MONTH_COLUMNS = [
        'Apr' => 2,
        'May' => 4,
        'Jun' => 6,
        'Jul' => 8,
        'Aug' => 10,
        'Sep' => 12,
        'Oct' => 14,
        'Nov' => 16,
        'Dec' => 18,
        'Jan' => 20,
        'Feb' => 22,
        'Mar' => 24,
    ];

    public function __construct()
    {
        $this->allValueTotal       = 0;
        $this->allDaysTotal        = 0;
        $this->monthsSubheadLabels = ['Booked by', 'Nightly Rate'];

        $this->borderStyle = [
            'borders' => [
                'outline' => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];

        $this->numberCurrencyStyle = [
            'numberFormat' => [
                'formatCode' => NumberFormat::FORMAT_CURRENCY_USD,
            ],
        ];

        $this->totalStyles = [
            'font' => ['bold' => true],
            'borders' => [
                'top' => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => ['argb' => 'FF000000'],
                ],
                'bottom' => [
                    'borderStyle' => Border::BORDER_MEDIUM,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
        ];

        $this->headingStyles = [
            'font' => ['bold' => true],
        ];
    }

    private function getMonthNames(): DatePeriod
    {
        $taxYearStart = new DateTimeImmutable('1st April');

        return new DatePeriod(
            $taxYearStart,
            DateInterval::createFromDateString('1 month'),
            $taxYearStart->add(DateInterval::createFromDateString('1 year')),
        );
    }

    /**
     * Gets the start column for specific month.
     *
     * @param string $months
     *
     * @throws Exception
     **/
    private function getMonthStartColumn(string $month): int
    {
        if (! isset(self::MONTH_COLUMNS[$month])) {
            throw new Exception('Month not found');
        }

        return self::MONTH_COLUMNS[$month];
    }

    private function addHeaderLabels(): void
    {
        for ($i = 1; $i <= self::DAYS_IN_MONTH; $i++) {
            $this->activeWorksheet->setCellValue([1, $i + self::DATE_START_ROW_NUMBER + 1], $i);
        }

        $this->activeWorksheet->setCellValue([1, self::DATE_START_ROW_NUMBER + self::TOTAL_ROW_NUMBER + 1], 'Total');
        $this->activeWorksheet->getStyle(1)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    }

    /**
     * @param array<string, array<float, string>> $calendar
     *
     * @return array<string, array<float, string>>
     */
    private function addMonthCols(iterable $calendar): array
    {
        $monthsNames = $this->getMonthNames();

        foreach ($monthsNames as $position => $monthName) {
            $formattedMonthName = $monthName->format('M');
            $monthStartColumn   = $this->getMonthStartColumn($formattedMonthName);
            $this->activeWorksheet->setCellValue([$monthStartColumn, self::DATE_START_ROW_NUMBER], $formattedMonthName);

            $cellAddressLeft  = $this->activeWorksheet
                ->getCell([$monthStartColumn, self::DATE_START_ROW_NUMBER])
                ->getCoordinate();
            $cellAddressRight = $this->activeWorksheet
                ->getCell([$monthStartColumn + 1, self::DATE_START_ROW_NUMBER])
                ->getCoordinate();
            $monthCellAddress = $cellAddressLeft . ':' . $cellAddressRight;

            $this->activeWorksheet
                ->mergeCells($monthCellAddress);
            $this->activeWorksheet
                ->getStyle($monthCellAddress)
                ->getAlignment()
                ->setHorizontal(Alignment::HORIZONTAL_CENTER);

            $monthBlockHeadAddress =
                $cellAddressLeft . ':' . substr(
                    $cellAddressRight,
                    0,
                    strlen($cellAddressRight) - 1,
                ) . self::MONTH_BLOCK_HEAD_ROWS;
            $this->activeWorksheet
                ->getStyle($monthBlockHeadAddress)
                ->applyFromArray($this->borderStyle);

            $monthBlockAddress =
                $cellAddressLeft . ':' . substr(
                    $cellAddressRight,
                    0,
                    strlen($cellAddressRight) - 1,
                ) . self::MONTH_BLOCK_ROWS;
            $this->activeWorksheet
                ->getStyle($monthBlockAddress)
                ->applyFromArray($this->borderStyle);

            $this->setDailyValue($monthStartColumn, 'Booked By');
            $this->setDailyValue($monthStartColumn + 1, 'Nightly Rate');

            $monthData = array_filter(
                $calendar,
                static function ($key) use ($monthName) {
                    $dateTime = DateTimeImmutable::createFromFormat('Y-m-d', $key);

                    return $dateTime->format('M') === $monthName->format('M');
                },
                ARRAY_FILTER_USE_KEY,
            );

            $monthTotal = 0;
            foreach ($monthData as $record) {
                if (! array_key_exists('nightly_rate', $record)) {
                    continue;
                }

                $monthTotal += (float) $record['nightly_rate'];
            }

            for ($day = 1; $day <= self::DAYS_IN_MONTH; $day++) {
                $monthDayData = array_filter(
                    $monthData,
                    static function ($key) use ($day) {
                        $dateTime    = new DateTimeImmutable($key);
                        $dateTimeDay = (int) ($dateTime->format('j'));

                        return $dateTimeDay === (int) $day;
                    },
                    ARRAY_FILTER_USE_KEY,
                );

                $dayRow = self::DATE_START_ROW_NUMBER + $day + 1;

                if (count($monthDayData) <= 0) {
                    continue;
                }

                $data        = array_shift($monthDayData);
                $nightlyRate = isset($data['nightly_rate']) ? number_format(
                    $data['nightly_rate'],
                    2,
                    '.',
                    '',
                ) : null;
                $bookingType = $data['booking_type'] ?? null;

                $this->activeWorksheet->setCellValueExplicit(
                    [$monthStartColumn, $dayRow],
                    $bookingType,
                    DataType::TYPE_STRING,
                );
                $this->activeWorksheet->setCellValueExplicit(
                    [$monthStartColumn + 1, $dayRow],
                    $nightlyRate,
                    DataType::TYPE_NUMERIC,
                );

                $this->activeWorksheet->getStyle([$monthStartColumn + 1, $dayRow])
                    ->getNumberFormat()
                    ->setFormatCode('0.00');
            }

            $this->activeWorksheet->setCellValue([$monthStartColumn + 1, self::MONTH_BLOCK_ROWS], $monthTotal);

            $this->activeWorksheet
                ->getStyle(self::MONTH_BLOCK_ROWS)
                ->applyFromArray($this->numberCurrencyStyle);
        }

        return $calendar;
    }

    private function setDailyValue(int $startColumn, string $type): void
    {
        $this->activeWorksheet
            ->setCellValue([$startColumn, self::DATE_START_ROW_NUMBER + 1], $type);
        $columnCoordinate = $this->activeWorksheet
            ->getCell([$startColumn, self::DATE_START_ROW_NUMBER + 1])
            ->getCoordinate();
        $columnCoordinate = substr($columnCoordinate, 0, strlen($columnCoordinate) - 1);
        $this->activeWorksheet
            ->getColumnDimension($columnCoordinate)
            ->setAutoSize(true);
    }

    /** @param array $monthData */
    private function dailyData(array $monthData, int $monthStartColumn): void
    {
        for ($day = 1; $day <= self::DAYS_IN_MONTH; $day++) {
            $dayRow  = self::DATE_START_ROW_NUMBER + $day + 1;
            $dayData = $monthData[date('Y-m') . '-' . sprintf('%02d', $day)] ?? [];

            if (empty($dayData)) {
                continue;
            }

            $data        = reset($dayData);
            $nightlyRate = $this->toCents($data[self::NIGHTLY_RATE]) ?? null;
            $bookingType = $data[self::BOOKING_TYPE] ?? null;

            $this->activeWorksheet->setCellValue([$monthStartColumn, $dayRow], $bookingType);
            $this->activeWorksheet->setCellValue([$monthStartColumn + 1, $dayRow], $nightlyRate);
        }
    }

    private function addSummary(): int
    {
        $subCategoryCount = 0;
        foreach ($this->categories as $subCategories) {
            $subCategoryCount += count($subCategories);
        }

        $numberOfSummaryRows = count($this->categories) * self::SUMMARY_ROW_COUNT_PER_CATEGORY
            + $subCategoryCount
            + self::CATEGORY_OFFSET;
        $summaryBoxAddress   =
            'B' . self::SUMMARY_START_ROW_NUMBER . ':E' . self::SUMMARY_START_ROW_NUMBER + $numberOfSummaryRows;
        $this->activeWorksheet
            ->getStyle($summaryBoxAddress)
            ->applyFromArray($this->borderStyle);

        $summaryBoxTitleAddress = 'B' . self::SUMMARY_START_ROW_NUMBER . ':E' . self::SUMMARY_START_ROW_NUMBER;
        $this->activeWorksheet
            ->mergeCells($summaryBoxTitleAddress);
        $this->activeWorksheet
            ->setCellValue(
                [self::SUMMARY_CATEGORY_NAME_COLUMN, self::SUMMARY_START_ROW_NUMBER],
                'Summary of Use',
            );
        $this->activeWorksheet
            ->getStyle(self::SUMMARY_START_ROW_NUMBER)
            ->getAlignment()
            ->setHorizontal(Alignment::HORIZONTAL_CENTER);

        $this->activeWorksheet
            ->getStyle(self::SUMMARY_START_ROW_NUMBER)
            ->applyFromArray($this->headingStyles);

        return $numberOfSummaryRows;
    }

    /**
     * @param array $calendar
     *
     * @return array
     */
    private function addCategoryInfo(array $calendar): array
    {
        $categoryIndex  = 0;
        $categoryOffset = self::CATEGORY_OFFSET;

        foreach ($this->categories as $categoryName => $subCategories) {
            $categoryStartRow = self::SUMMARY_START_ROW_NUMBER + $categoryOffset + 1;

            $this->activeWorksheet
                ->setCellValue([self::SUMMARY_CATEGORY_NAME_COLUMN, $categoryStartRow], $categoryName);
            $this->activeWorksheet
                ->getStyle($categoryStartRow)
                ->applyFromArray($this->headingStyles);

            $this->activeWorksheet
                ->setCellValue([
                    self::SUMMARY_DAYS_COLUMN,
                    self::SUMMARY_START_ROW_NUMBER,
                ], 'Days');
            $this->activeWorksheet
                ->getStyle(self::SUMMARY_START_ROW_NUMBER)
                ->applyFromArray($this->headingStyles);
            $this->activeWorksheet
                ->setCellValue([self::SUMMARY_VALUE_COLUMN, self::SUMMARY_START_ROW_NUMBER], 'Value $');

            $this->activeWorksheet
                ->getStyle(self::SUMMARY_START_ROW_NUMBER + $categoryIndex)
                ->applyFromArray($this->headingStyles);

            $categoryTotalDays  = 0;
            $categoryTotalValue = 0;

            foreach ($subCategories as $position => $subCategory) {
                $subCategoryRowNumber = $categoryStartRow + $position + 1;

                $categoryStat = $this->getCategoryStat($calendar, $subCategory);

                $categoryTotalDays  .= $categoryStat['days'];
                $categoryTotalValue .= $categoryStat['total'];

                $this->activeWorksheet->setCellValue(
                    [self::SUMMARY_CATEGORY_NAME_COLUMN, $subCategoryRowNumber],
                    $subCategory,
                );
                $this->activeWorksheet->setCellValue([
                    self::SUMMARY_DAYS_COLUMN,
                    $subCategoryRowNumber,
                ], $categoryStat['days']);
                $this->activeWorksheet->setCellValue(
                    [self::SUMMARY_VALUE_COLUMN, $subCategoryRowNumber],
                    $categoryStat['total'],
                );

                $this->activeWorksheet
                    ->getStyle('E' . $subCategoryRowNumber)
                    ->applyFromArray($this->numberCurrencyStyle);
            }

            $this->allValueTotal += $categoryTotalValue;
            $this->allDaysTotal  += $categoryTotalDays;

            $this->activeWorksheet
                ->setCellValue(
                    [self::SUMMARY_TOTALS_COLUMN, $subCategoryRowNumber + 1],
                    $categoryName . ' total',
                );

            $this->activeWorksheet
                ->getStyle($subCategoryRowNumber + 1)->applyFromArray($this->headingStyles);

            $this->activeWorksheet
                ->setCellValue([
                    self::SUMMARY_DAYS_COLUMN,
                    $subCategoryRowNumber + 1,
                ], $categoryTotalDays);
            $this->activeWorksheet
                ->getStyle($subCategoryRowNumber + 1)->applyFromArray($this->headingStyles);

            $this->activeWorksheet
                ->setCellValue([self::SUMMARY_VALUE_COLUMN, $subCategoryRowNumber + 1], $categoryTotalValue);

            $this->activeWorksheet
                ->getStyle($subCategoryRowNumber + 1)
                ->applyFromArray($this->headingStyles);
            $this->activeWorksheet
                ->getStyle('E' . $subCategoryRowNumber + 1)
                ->applyFromArray($this->numberCurrencyStyle);

            $categoryOffset += count($subCategories) + 1;
        }

        return $calendar;
    }

    private function addTotals(int $numberOfSummaryRows): void
    {
        $propertyTotalRowNumber = self::SUMMARY_START_ROW_NUMBER + $numberOfSummaryRows;
        $this->activeWorksheet
            ->setCellValue([self::SUMMARY_TOTALS_COLUMN, $propertyTotalRowNumber], 'Property Total');
        $this->activeWorksheet
            ->getStyle($propertyTotalRowNumber)->applyFromArray($this->headingStyles);
        $this->activeWorksheet
            ->setCellValue([
                self::SUMMARY_DAYS_COLUMN,
                $propertyTotalRowNumber,
            ], $this->allDaysTotal);
        $this->activeWorksheet
            ->setCellValue([self::SUMMARY_VALUE_COLUMN, $propertyTotalRowNumber], $this->allValueTotal);
        $this->activeWorksheet
            ->getStyle('E' . $propertyTotalRowNumber)->applyFromArray($this->numberCurrencyStyle);
    }

    /**
     * @param string[] $calendar
     *
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function write(int $propertyId, array $calendar): void
    {
        $this->spreadsheet     = new Spreadsheet();
        $this->activeWorksheet = $this->spreadsheet->getActiveSheet();

        $this->addHeaderLabels();
        $calendar            = $this->addMonthCols($calendar);
        $this->categories    = $this->getUniqueBookingTypes([$calendar]);
        $numberOfSummaryRows = $this->addSummary();
        $calendar            = $this->addCategoryInfo($calendar);
        $this->addTotals($numberOfSummaryRows);

        $writer = new Xlsx($this->spreadsheet);
        $writer->save('/tmp/' . $propertyId . '.xlsx');
    }

    /**
     * @param array $data
     *
     * @return array
     */
    private function getUniqueBookingTypes(array $data): array
    {
        $calendar = [];
        foreach ($data as $column) {
            $calendar[] =  array_column($column, 'booking_type');
        }

        $uniqueTypes = [
            'Bachcare Date' => [],
            'Owner Date' => [],
        ];

        $uniqueBookingType = array_unique($calendar[0]);
        foreach ($uniqueBookingType as $type) {
            $types = [];
            if (str_contains(strtolower($type), 'bachcare')) {
                $uniqueTypes['Bachcare Date'][] = $type;
            }

            if (
                ! str_contains(strtolower($type), 'own')
                && ! str_contains(strtolower($type), 'house')
                && (! str_contains(strtolower($type), 'term'))
            ) {
                continue;
            }

            $uniqueTypes['Owner Date'][] = $type;
        }

        return $uniqueTypes;
    }

    /**
     * Calculates total value in a list of record
     *
     * @param string[] $data
     */
    private function getDataTotal(array $data): float
    {
        $total = 0;

        foreach ($data as $record) {
            $total += (float) $record['nightly_rate'];
        }

        return $total;
    }

    /**
     * Calculates total days and value for a specific category
     *
     * @param string[] $data
     *
     * @return string[]
     */
    private function getCategoryStat(array $data, string $category): array
    {
        $categoryData = array_filter(
            $data,
            static function ($record) use ($category) {
                return array_key_exists('booking_type', (array) $record) && $record['booking_type'] === $category;
            },
        );

        $categoryTotal = $this->getDataTotal($categoryData);
        $categoryDays  = count($categoryData);

        return ['total' => $categoryTotal, 'days' => $categoryDays];
    }
}
