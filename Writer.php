<?php

declare(strict_types=1);

namespace SykesCottages\NewZealandOwnerTaxService\Writer;

interface Writer
{
    /** @param array<string, array<float, string>> $calendar */
    public function write(int $propertyId, array $calendar): void;
}
