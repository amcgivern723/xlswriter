<?php

declare(strict_types=1);

namespace SykesCottages\NewZealandOwnerTaxService\Acceptance\Writer;

use Generator;
use PHPUnit\Framework\TestCase;
use SykesCottages\NewZealandOwnerTaxService\Writer\CsvWriter;
use SykesCottages\NewZealandOwnerTaxService\Writer\Writer;
use SykesCottages\NewZealandOwnerTaxService\Writer\XlsWriter;
use Throwable;

use function abs;
use function file_exists;
use function file_get_contents;
use function filesize;
use function json_decode;
use function unlink;

class WriterTest extends TestCase
{
    private const DIRECTORY = __DIR__ . '/../Fixture/';

    /**
     * @param array<string, array{nightly_rate: float, booking_type: string}> $calendar
     *
     * @dataProvider csvUseCaseGenerator
     */
    public function testCalendarCanBeWrittenToTmpDirectory(
        int $propertyId,
        array $calendar,
        string $expectedFileContents,
        string $fileExtension,
        Writer $writer,
    ): void {
        $fileName = '/tmp/' . $propertyId . $fileExtension;

        try {
            $writer->write($propertyId, $calendar);
            self::assertFileExists($fileName);
            self::assertStringEqualsFile($fileName, $expectedFileContents);
        } catch (Throwable $exception) {
            self::fail($exception->getMessage());
        } finally {
            if (file_exists($fileName)) {
                unlink($fileName);
            }
        }
    }

    public function csvUseCaseGenerator(): Generator
    {
        yield 'it writes to csv' => [
            12345,
            [
                '2021-04-01' => [
                    'nightly_rate' => 0.00,
                    'booking_type' => 'Owner',
                ],
                '2021-04-02' => [],
                '2021-04-03' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
            ],
            "date,nightly_rate,booking_type\n2021-04-01,0,Owner\n2021-04-02\n2021-04-03,420.69,Customer\n",
            '.csv',
            new CsvWriter(),
        ];
    }

    /** @dataProvider xlsUseCaseGenerator */
    public function testXlsCalendarCanBeWrittenToTmpDirectory(
        int $propertyId,
        string $fixtureFileName,
        string $expectedFilename,
    ): void {
        $calendar = json_decode(file_get_contents($fixtureFileName), true);
        $writer   =  new XlsWriter();

        $filename = '/tmp/' . $propertyId . '.xlsx';
        try {
            $writer->write($propertyId, $calendar);
            $lengthOfExpectedFileContents = filesize($expectedFilename);
            $lengthOfFileContents         = filesize($filename);

            $fileSizeDiff = abs($lengthOfExpectedFileContents - $lengthOfFileContents);

            //THIS IS TO ALLOW FOR 1 byte differences between runs
            //probably caused by OS specific file system differences
            self::assertLessThanOrEqual(1, $fileSizeDiff);
        } catch (Throwable $exception) {
            self::fail($exception->getMessage());
        } finally {
            if (file_exists($filename)) {
                unlink($filename);
            }
        }
    }

    public function xlsUseCaseGenerator(): Generator
    {
        yield 'it writes to xlsx' => [
            12345,
            self::DIRECTORY . 'ownerExampleDataForOneMonth.json',
            self::DIRECTORY . '12345.xlsx',
        ];

        yield 'it writes to xlsx version 2' => [
            6789,
            self::DIRECTORY . 'ownerFullTaxYearExampleData.json',
            self::DIRECTORY . '6789.xlsx',
            new XlsWriter(),
        ];
    }
}
