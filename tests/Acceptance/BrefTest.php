<?php

declare(strict_types=1);

namespace SykesCottages\NewZealandOwnerTaxService\Acceptance;

use Bref\Context\Context;
use Bref\Logger\StderrLogger;
use PHPUnit\Framework\MockObject\Exception;
use PHPUnit\Framework\TestCase;
use SykesCottages\NewZealandOwnerTaxService\Client\HttpClient;
use SykesCottages\NewZealandOwnerTaxService\Controllers\HandlerController;
use SykesCottages\NewZealandOwnerTaxService\Exceptions\MissingEvent;
use SykesCottages\NewZealandOwnerTaxService\Exceptions\UnrecognisedTask;

use function file_get_contents;
use function json_decode;

class BrefTest extends TestCase
{
    public const MISSING_EVENT_JSON     = '/Json/missingEvent.json';
    public const UNRECOGNISED_TASK_JSON = '/Json/unrecognisedTask.json';
    public const PING_TASK_JSON         = '/Json/pingTask.json';

    private function getEvent(string $file): array
    {
        return json_decode(
            file_get_contents(__DIR__ . $file),
            true,
        );
    }

    private function getContext(): Context
    {
        return new Context('', 300, '', '');
    }

    /** @throws Exception */
    private function getHandler(): HandlerController
    {
        return new HandlerController(
            $this->createMock(StderrLogger::class),
            $this->createMock(HttpClient::class),
        );
    }

    /**
     * @throws Exception
     * @throws MissingEvent
     * @throws UnrecognisedTask
     */
    public function testMissingEventReturnsException(): void
    {
        $this->expectException(MissingEvent::class);
        $this->getHandler()->handle(
            $this->getEvent(self::MISSING_EVENT_JSON),
            $this->getContext(),
        );
    }

    /**
     * @throws Exception
     * @throws MissingEvent
     * @throws UnrecognisedTask
     */
    public function testUnrecognisedTaskReturnsException(): void
    {
        $this->expectException(UnrecognisedTask::class);
        $this->getHandler()->handle(
            $this->getEvent(self::UNRECOGNISED_TASK_JSON),
            $this->getContext(),
        );
    }
}
