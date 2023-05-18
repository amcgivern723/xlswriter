<?php

declare(strict_types=1);

namespace SykesCottages\NewZealandOwnerTaxService\Unit\Processor;

use DateTimeImmutable;
use DateTimeInterface;
use Exception;
use Generator;
use PHPUnit\Framework\TestCase;
use SykesCottages\NewZealandOwnerTaxService\Processor\Calendar;

class CalendarTest extends TestCase
{
    /**
     * @param array<string[]>                     $propertyData
     * @param array<string, array<float, string>> $expected
     *
     * @throws Exception
     *
     * @dataProvider useCaseGenerator
     */
    public function testCalendar(
        array $propertyData,
        DateTimeInterface $taxYearToReportOn,
        array $expected,
    ): void {
        $sut = new Calendar(
            $propertyData,
            $taxYearToReportOn,
        );

        $calendarData = $sut->generate();

        foreach ($expected as $date => $value) {
            self::assertEquals($value, $calendarData[$date]);
        }
    }

    public function useCaseGenerator(): Generator
    {
        yield 'it returns data in the correct format for a single day' => [
            [
                [
                    'nightly_rate' => 0.00,
                    'start_date' => '2021-04-07',
                    'end_date' => '2021-04-08',
                    'booking_type' => 'Owner',
                ],
            ],
            DateTimeImmutable::createFromFormat('Y', '2021'),
            [
                '2021-04-07' => [
                    'nightly_rate' => 0.00,
                    'booking_type' => 'Owner',
                ],
            ],
        ];

        yield 'it returns an empty array for days with no bookings' => [
            [
                [
                    'nightly_rate' => 0.00,
                    'start_date' => '2021-04-07',
                    'end_date' => '2021-04-08',
                    'booking_type' => 'Owner',
                ],
            ],
            DateTimeImmutable::createFromFormat('Y', '2021'),
            [
                '2021-10-10' => [],
            ],
        ];

        yield 'it can run across a date range and return either an empty result or a data set' => [
            [
                [
                    'nightly_rate' => 0.00,
                    'start_date' => '2021-10-09',
                    'end_date' => '2021-10-10',
                    'booking_type' => 'Owner',
                ],
                [
                    'nightly_rate' => 0.00,
                    'start_date' => '2021-10-11',
                    'end_date' => '2021-10-12',
                    'booking_type' => 'Owner',
                ],
            ],
            DateTimeImmutable::createFromFormat('Y', '2021'),
            [
                '2021-10-09' => [
                    'nightly_rate' => 0.00,
                    'booking_type' => 'Owner',
                ],
                '2021-10-10' => [],
                '2021-10-11' => [
                    'nightly_rate' => 0.00,
                    'booking_type' => 'Owner',
                ],
            ],
        ];

        yield 'it can handle a booking of multiple days' => [
            [
                [
                    'nightly_rate' => 420.69,
                    'start_date' => '2021-06-30',
                    'end_date' => '2021-07-05',
                    'booking_type' => 'Customer',
                ],
            ],
            DateTimeImmutable::createFromFormat('Y', '2021'),
            [
                '2021-06-29' => [],
                '2021-06-30' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-01' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-02' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-03' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-04' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-05' => [],
            ],
        ];

        yield 'it can handle a back to back bookings' => [
            [
                [
                    'nightly_rate' => 420.69,
                    'start_date' => '2021-06-30',
                    'end_date' => '2021-07-03',
                    'booking_type' => 'Customer',
                ],
                [
                    'nightly_rate' => 0.00,
                    'start_date' => '2021-07-03',
                    'end_date' => '2021-07-06',
                    'booking_type' => 'Owner',
                ],
            ],
            DateTimeImmutable::createFromFormat('Y', '2021'),
            [
                '2021-06-29' => [],
                '2021-06-30' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-01' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-02' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-07-03' => [
                    'nightly_rate' => 0.00,
                    'booking_type' => 'Owner',
                ],
                '2021-07-04' => [
                    'nightly_rate' => 0.00,
                    'booking_type' => 'Owner',
                ],
                '2021-07-05' => [
                    'nightly_rate' => 0.00,
                    'booking_type' => 'Owner',
                ],
                '2021-07-06' => [],
            ],
        ];

        yield 'it can handle bookings in leap years' => [
            [
                [
                    'nightly_rate' => 420.69,
                    'start_date' => '2020-02-27',
                    'end_date' => '2020-03-03',
                    'booking_type' => 'Customer',
                ],
            ],
            DateTimeImmutable::createFromFormat('Y', '2019'),
            [
                '2020-02-26' => [],
                '2020-02-27' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2020-02-28' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2020-02-29' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2020-03-01' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2020-03-02' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2020-03-03' => [],
            ],
        ];

        yield 'it can handle bookings in feb/mar crossover in non-leap years' => [
            [
                [
                    'nightly_rate' => 420.69,
                    'start_date' => '2021-02-27',
                    'end_date' => '2021-03-03',
                    'booking_type' => 'Customer',
                ],
            ],
            DateTimeImmutable::createFromFormat('Y', '2020'),
            [
                '2021-02-26' => [],
                '2021-02-27' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-02-28' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-03-01' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-03-02' => [
                    'nightly_rate' => 420.69,
                    'booking_type' => 'Customer',
                ],
                '2021-03-03' => [],
            ],
        ];
    }
}
