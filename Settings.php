<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin;

use Aurora\System\SettingsProperty;

/**
 * @property bool $Disabled
 * @property bool $AllowAppointments
 */

class Settings extends \Aurora\System\Module\Settings
{
    protected function initDefaults()
    {
        $this->aContainer = [
            "Disabled" => new SettingsProperty(
                false,
                "bool",
                null,
                "Setting to true disables the module",
            ),
            "AllowAppointments" => new SettingsProperty(
                true,
                "bool",
                null,
                "If true, allows for adding event attendees and enables processing event invitations in mail",
            ),
        ];
    }
}
