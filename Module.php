<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin;

use Aurora\Modules\Core\Module as CoreModule;
use Aurora\Modules\Calendar\Module as CalendarModule;
use Aurora\Modules\Calendar\Classes\Helper as CalendarHelper;
use Aurora\System\Api;
use MailSo\Mime\Email;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2023, Afterlogic Corp.
 *
 * @property Settings $oModuleSettings
 *
 * @package Modules
 */
class Module extends \Aurora\System\Module\AbstractModule
{
    /** @var Manager */
    protected $oManager = null;

    public $oApiFileCache = null;
    //	public $oApiCalendarDecorator = null;
    //	public $oApiUsersManager = null;

    /**
     * @return Module
     */
    public static function getInstance()
    {
        return parent::getInstance();
    }

    /**
     * @return Module
     */
    public static function Decorator()
    {
        return parent::Decorator();
    }

    /**
     * @return Settings
     */
    public function getModuleSettings()
    {
        return $this->oModuleSettings;
    }

    public function getManager()
    {
        if ($this->oManager === null) {
            $this->oManager = new Manager($this);
        }

        return $this->oManager;
    }

    public function getCacheManager()
    {
        if ($this->oApiFileCache === null) {
            $this->oApiFileCache = new \Aurora\System\Managers\Filecache();
        }

        return $this->oApiFileCache;
    }

    public function init()
    {
        $this->AddEntries([
            'invite' => 'EntryInvite'
        ]);

        $this->subscribeEvent('Calendar::CreateIcs', array($this, 'onCreateIcs'));
        $this->subscribeEvent('Calendar::populateVCalendar', array($this, 'onPopulateVCalendar'));
        $this->subscribeEvent('Calendar::DeleteEvent', array($this, 'onDeleteEvent'));
        $this->subscribeEvent('Calendar::UpdateEventAttendees', array($this, 'onUpdateEventAttendees'));
        $this->subscribeEvent('CalendarMeetingsPlugin::processICS::UpdateEvent', array($this, 'onProcessICSUpdateEvent'));
        $this->subscribeEvent('CalendarMeetingsPlugin::processICS::Cancel', array($this, 'onProcessICSCancel'));
        $this->subscribeEvent('Calendar::processICS::UpdateEvent', array($this, 'onProcessICSUpdateEvent'));
        $this->subscribeEvent('Calendar::processICS::Cancel', array($this, 'onProcessICSCancel'));
        $this->subscribeEvent('Calendar::processICS::AddAttendeesToResult', array($this, 'onAddAttendeesToResult'));
        $this->subscribeEvent('Calendar::parseEvent', array($this, 'onParseEvent'));

        $this->aErrors = [
            Enums\ErrorCodes::CannotSendAppointmentMessage => $this->i18N('ERROR_CANNOT_SEND_APPOINTMENT_MESSAGE'),
            Enums\ErrorCodes::CannotSendAppointmentMessageNoOrganizer => $this->i18N('ERROR_CANNOT_SEND_APPOINTMENT_MESSAGE_NO_ORGANIZER')
        ];
    }

    /**
     * @param int $UserId
     * @param string $File
     * @param string $FromEmail
     * @return boolean
     * @throws \Aurora\System\Exceptions\ApiException
     */
    public function UpdateAttendeeStatus($UserId, $File, $FromEmail)
    {
        \Aurora\System\Api::CheckAccess($UserId);
        $sUserPublicId = \Aurora\System\Api::getUserPublicIdById($UserId);
        $mResult = false;

        if (empty($File) || empty($FromEmail)) {
            throw new \Aurora\System\Exceptions\ApiException(\Aurora\System\Notifications::InvalidInputParameter);
        }
        $sData = $this->getCacheManager()->get($sUserPublicId, $File, '', CalendarModule::GetName());
        if (!empty($sData)) {
            $mResult = $this->getManager()->processICS($sUserPublicId, $sData, $FromEmail, true);
        }

        return $mResult;
    }

    /**
     *
     * @param int $UserId
     * @param string $CalendarId
     * @param string $EventId
     * @param string $File
     * @param string $AppointmentAction
     * @param string $Attendee
     * @return array|boolean
     * @throws \Aurora\System\Exceptions\ApiException
     */
    public function SetAppointmentAction($UserId, $CalendarId, $EventId, $File, $AppointmentAction, $Attendee, $AllEvents = 2, $RecurrenceId = null)
    {
        \Aurora\System\Api::CheckAccess($UserId);
        $sUserPublicId = \Aurora\System\Api::getUserPublicIdById($UserId);
        $mResult = false;

        if (empty($AppointmentAction) || empty($CalendarId)) {
            throw new \Aurora\System\Exceptions\ApiException(\Aurora\System\Notifications::InvalidInputParameter);
        }
        $sData = '';
        if (!empty($EventId)) {
            $aEventData =  $this->getManager()->getEvent($sUserPublicId, $CalendarId, $EventId);
            if (isset($aEventData) && isset($aEventData['vcal']) && $aEventData['vcal'] instanceof \Sabre\VObject\Component\VCalendar) {
                $oVCal = $aEventData['vcal'];
                $oVCal->METHOD = 'REQUEST';
                $sData = $oVCal->serialize();
            }
        } elseif (!empty($File)) {
            $sData = $this->getCacheManager()->get($sUserPublicId, $File, '', CalendarModule::GetName());
        }
        if (!empty($sData)) {
            $mProcessResult = $this->getManager()->appointmentAction($sUserPublicId, $Attendee, $AppointmentAction, $CalendarId, $sData, $AllEvents, $RecurrenceId);
            if ($mProcessResult) {
                $mResult = array(
                    'Uid' => $mProcessResult
                );
            }
        }

        return $mResult;
    }

    public function EntryInvite()
    {
        $sResult = '';
        $aInviteValues = \Aurora\System\Api::DecodeKeyValues($this->oHttp->GetQuery('invite'));

        if (isset($aInviteValues['organizer'])) {
            $sOrganizerPublicId = $aInviteValues['organizer'];
            $oOrganizerUser = CoreModule::Decorator()->GetUserByPublicId($sOrganizerPublicId);
            if ($sOrganizerPublicId && isset($aInviteValues['attendee'], $aInviteValues['eventId'], $aInviteValues['action'])) {
                $calendarId = $aInviteValues['calendarId'];
                if (empty($calendarId)) {
                    $calendarId = $this->getManager()->findEventInCalendars($sOrganizerPublicId, $aInviteValues['eventId']);
                }
                $sAction = '';
                $oEvent = $this->getManager()->getEvent($sOrganizerPublicId, $calendarId, $aInviteValues['eventId']);
                $oModuleManager = \Aurora\System\Api::GetModuleManager();
                $sTheme = $oModuleManager->getModuleConfigValue('CoreWebclient', 'Theme');
                if ($oEvent && is_array($oEvent) && 0 < count($oEvent) && isset($oEvent[0])) {
                    $sResult = \file_get_contents($this->GetPath() . '/templates/CalendarEventInviteExternal.html');
                    if (is_string($sResult)) {
                        $dt = new \DateTime();
                        $dt->setTimestamp($oEvent[0]['startTS']);
                        if (!$oEvent[0]['allDay']) {
                            $sDefaultTimeZone = new \DateTimeZone($oOrganizerUser->DefaultTimeZone);
                            $dt->setTimezone($sDefaultTimeZone);
                        }

                        $sAction = $aInviteValues['action'];
                        $sActionColor = 'green';
                        $sActionText = '';
                        switch (strtoupper($sAction)) {
                            case 'ACCEPTED':
                                $sActionColor = 'green';
                                $sActionText = 'Accepted';
                                break;
                            case 'DECLINED':
                                $sActionColor = 'red';
                                $sActionText = 'Declined';
                                break;
                            case 'TENTATIVE':
                                $sActionColor = '#A0A0A0';
                                $sActionText = 'Tentative';
                                break;
                        }

                        $sDateFormat = 'm/d/Y';
                        $sTimeFormat = 'h:i A';
                        switch ($oOrganizerUser->DateFormat) {
                            case \Aurora\System\Enums\DateFormat::DDMMYYYY:
                                $sDateFormat = 'd/m/Y';
                                break;
                            case \Aurora\System\Enums\DateFormat::DD_MONTH_YYYY:
                                $sDateFormat = 'd/m/Y';
                                break;
                            default:
                                $sDateFormat = 'm/d/Y';
                                break;
                        }
                        switch ($oOrganizerUser->TimeFormat) {
                            case \Aurora\System\Enums\TimeFormat::F24:
                                $sTimeFormat = 'H:i';
                                break;
                            case \Aurora\System\Enums\DateFormat::DD_MONTH_YYYY:
                                \Aurora\System\Enums\TimeFormat::F12;
                                $sTimeFormat = 'h:i A';
                                break;
                            default:
                                $sTimeFormat = 'h:i A';
                                break;
                        }
                        $sDateTime = $dt->format($sDateFormat . ' ' . $sTimeFormat);

                        $mResult = array(
                            '{{EVENT_NAME}}' => $oEvent[0]['subject'],
                            '{{EVENT_BEGIN}}' => ucfirst($this->i18N('EVENT_BEGIN')),
                            '{{EVENT_DATE}}' => $sDateTime,
                            '{{CALENDAR}}' => ucfirst($this->i18N('CALENDAR')),
                            '{{EVENT_DESCRIPTION}}' => $oEvent[0]['description'],
                            '{{EVENT_ACTION}}' => $sActionText,
                            '{{ACTION_COLOR}}' => $sActionColor,
                            '{{THEME_NAME}}' => $sTheme,
                        );

                        $sResult = strtr($sResult, $mResult);
                    } else {
                        $sResult = "Error occured";
                        \Aurora\System\Api::Log('Empty page template.', \Aurora\System\Enums\LogLevel::Error);
                    }
                } else {
                    \Aurora\System\Api::Log('Event not found.', \Aurora\System\Enums\LogLevel::Error);

                    $sResult = file_get_contents($this->GetPath() . '/templates/EventNotFound.html');

                    if (is_string($sResult)) {
                        $mResult = array(
                            '{{INFO}}' => ucfirst($this->i18N('ERROR_APPOINTMENT_NOT_FOUND')),
                            '{{THEME_NAME}}' => $sTheme,
                        );

                        $sResult = strtr($sResult, $mResult);
                    } else {
                        $sResult = "Error occured";
                        \Aurora\System\Api::Log('Empty page template.', \Aurora\System\Enums\LogLevel::Error);
                    }
                }
                $sAttendee = $aInviteValues['attendee'];
                if (!empty($sAttendee)) {
                    $oAttendee = Email::Parse($sAttendee);
                    $sAttendee = $oAttendee->GetEmail();

                    if (isset($oEvent) && isset($oEvent['vcal']) && $oEvent['vcal'] instanceof \Sabre\VObject\Component\VCalendar) {
                        $oVCal = $oEvent['vcal'];
                        $oVCal->METHOD = 'REQUEST';
                        $sData = $oVCal->serialize();
                        $oAttendeeUser = CoreModule::Decorator()->GetUserByPublicId($sAttendee);
                        $bIsExternalAttendee = false;
                        if ($oAttendeeUser) {
                            $sOrganizerPublicId = $oAttendeeUser->PublicId;
                        } else {
                            $bIsExternalAttendee = true;
                        }
                        $actionResult = false;
                        try {
                            $actionResult = $this->getManager()->appointmentAction($sOrganizerPublicId, $sAttendee, $sAction, $calendarId, $sData, 2, null, true, $bIsExternalAttendee);
                        } catch (\Aurora\System\Exceptions\Exception $e) {
                            Api::LogException($e);
                            $actionResult = false;
                        }
                        if (!$actionResult) {
                            $sResult = file_get_contents($this->GetPath() . '/templates/EventNotFound.html');

                            if (is_string($sResult)) {
                                $mResult = array(
                                    '{{INFO}}' => ucfirst($this->i18N('ERROR_APPOINTMENT_UPDATE_STATUS')),
                                    '{{THEME_NAME}}' => $sTheme,
                                );

                                $sResult = strtr($sResult, $mResult);
                            }
                        }
                    }
                }
            }
        }

        return $sResult;
    }

    /**
     *
     * @return array
     */
    public function GetSettings()
    {
        \Aurora\System\Api::checkUserRoleIsAtLeast(\Aurora\System\Enums\UserRole::Anonymous);

        $aSettings = array(
            'AllowAppointments' => $this->oModuleSettings->AllowAppointments
        );

        return $aSettings;
    }

    public function onCreateIcs($aData, &$oIcs)
    {
        $oIcs->Attendee = isset($aData['Attendee']) ? $aData['Attendee'] : null;
        $oIcs->Type = $aData['Action'];
    }

    public function onPopulateVCalendar(&$aData, &$oVEvent)
    {
        /** @var \Sabre\VObject\Component\VEvent $oVEvent */
        $oEvent = & $aData['oEvent'];

        $oUser = \Aurora\System\Api::getAuthenticatedUser();

        $aAttendees = [];
        $aAttendeeEmails = [];
        $aObjAttendees = [];
        if (isset($oVEvent->ORGANIZER)) {
            $sOwnerEmail = str_replace('mailto:', '', strtolower((string) $oVEvent->ORGANIZER));
            $iPos = strpos($sOwnerEmail, 'principals/');
            if ($iPos !== false) {
                $sOwnerEmail = \trim(substr($sOwnerEmail, $iPos + 11), '/');
            }
        }

        if (isset($oVEvent->ATTENDEE)) {
            $aAttendeeEmails = [];
            foreach ($oEvent->Attendees as $aItem) {
                $sStatus = '';
                switch ($aItem['status']) {
                    case \Aurora\Modules\Calendar\Enums\AttendeeStatus::Accepted:
                        $sStatus = 'ACCEPTED';
                        break;
                    case \Aurora\Modules\Calendar\Enums\AttendeeStatus::Declined:
                        $sStatus = 'DECLINED';
                        break;
                    case \Aurora\Modules\Calendar\Enums\AttendeeStatus::Tentative:
                        $sStatus = 'TENTATIVE';
                        break;
                    case \Aurora\Modules\Calendar\Enums\AttendeeStatus::Unknown:
                        $sStatus = 'NEEDS-ACTION';
                        break;
                }

                $aAttendeeEmails[strtolower($aItem['email'])] = $sStatus;
            }

            $aObjAttendees = $oVEvent->ATTENDEE;
            unset($oVEvent->ATTENDEE);
            foreach ($aObjAttendees as $oAttendee) {
                $sAttendee = str_replace('mailto:', '', strtolower((string)$oAttendee));
                $oPartstat = $oAttendee->offsetGet('PARTSTAT');
                if (in_array($sAttendee, array_keys($aAttendeeEmails))) {
                    if (isset($oPartstat) && (string)$oPartstat === $aAttendeeEmails[$sAttendee]) {
                        $oVEvent->add($oAttendee);
                        $aAttendees[] = $sAttendee;
                    }
                }
            }
        }

        if (isset($oEvent->Attendees) && count($oEvent->Attendees) > 0) {
            if (!isset($oVEvent->ORGANIZER)) {
                $oVEvent->ORGANIZER = 'mailto:' . $oUser->PublicId;
            }
            foreach ($oEvent->Attendees as $oAttendee) {
                if (!in_array($oAttendee['email'], $aAttendees)) {
                    $oVEvent->add(
                        'ATTENDEE',
                        'mailto:' . $oAttendee['email'],
                        array(
                            'CN' => !empty($oAttendee['name']) ? $oAttendee['name'] : $oAttendee['email'],
                            'RSVP' => 'TRUE'
                        )
                    );
                }
            }
        } else {
            unset($oVEvent->ORGANIZER);
        }
    }

    public function onDeleteEvent($aData, &$oVCal)
    {
        $sUserPublicId = $aData['sUserPublicId'];
        /** @var \Sabre\VObject\Component\VCalendar $oVCal */
        $iIndex = CalendarHelper::getBaseVComponentIndex($oVCal->VEVENT);
        if ($iIndex !== false) {
            /** @var \Sabre\VObject\Component\VEvent $oVEvent */
            $oVEvent = $oVCal->VEVENT[$iIndex];

            $sOrganizer = (isset($oVEvent->ORGANIZER)) ?
                    str_replace('mailto:', '', strtolower((string)$oVEvent->ORGANIZER)) : null;
            $iPos = false;
            if (!empty($sOrganizer)) {
                $iPos = strpos($sOrganizer, 'principals/');
            }

            if ($iPos !== false) {
                $sOrganizer = \trim(substr($sOrganizer, $iPos + 11), '/');
            }

            if (isset($sOrganizer)) {
                if ($sOrganizer === $sUserPublicId) {
                    $oDateTimeNow = new \DateTimeImmutable("now");
                    /** @var \Sabre\VObject\Property\ICalendar\DateTime $oDTSTART */
                    $oDTSTART = $oVEvent->DTSTART;
                    $oDateTimeEvent = $oDTSTART->getDateTime();
                    $oDateTimeRepeat = CalendarHelper::getNextRepeat($oDateTimeNow, $oVEvent);
                    $bRrule = isset($oVEvent->RRULE);
                    $bEventFore = $oDateTimeEvent ? $oDateTimeEvent > $oDateTimeNow : false;
                    $bNextRepeatFore = $oDateTimeRepeat ? $oDateTimeRepeat > $oDateTimeNow : false;

                    if (isset($oVEvent->ATTENDEE) && ($bRrule ? $bNextRepeatFore : $bEventFore)) {
                        foreach ($oVEvent->ATTENDEE as $oAttendee) {
                            $sEmail = str_replace('mailto:', '', strtolower((string)$oAttendee));

                            $oVCalResult = clone $oVCal;
                            $oVCalResult->METHOD = 'CANCEL';
                            $sSubject = (string)$oVEvent->SUMMARY . ': Canceled';

                            Classes\Helper::sendAppointmentMessage($sUserPublicId, $sEmail, $sSubject, $oVCalResult, 'REQUEST');
                            unset($oVCal->METHOD);
                        }
                    }
                }
            }
        }
    }

    public function onUpdateEventAttendees($aData, &$oEvent)
    {
        $aEventData =  $this->getManager()->getEvent($aData['UserPublicId'], $oEvent->IdCalendar, $oEvent->Id);
        if (isset($aEventData[0]) && isset($aEventData[0]['attendees'])) {
            $oldAttendees = $aEventData[0]['attendees'];
            $newAttendees =  @json_decode($aData['attendees'], true);

            $deleteAttendees = [];
            foreach ($oldAttendees as $old) {
                $found = false;
                foreach ($newAttendees as $new) {
                    if (strtolower($new['email']) === strtolower($old['email'])) {
                        $found = true;
                        break;
                    }
                }
                if (!$found) {
                    $deleteAttendees[] = $old;
                }
            }

            $oVCalResult = clone $aEventData['vcal'];
            $oVCalResult->METHOD = 'CANCEL';
            if (isset($oVCalResult->VEVENT)) {
                $sSubject = (string) $oVCalResult->VEVENT->SUMMARY . ': Canceled';
                $oVCalResult->VEVENT->SEQUENCE = (int) $oVCalResult->VEVENT->SEQUENCE->getValue() + 1;

                foreach ($deleteAttendees as $deleteAttendee) {
                    $sEmail = $deleteAttendee['email'];

                    unset($oVCalResult->VEVENT->ATTENDEE);

                    $oVCalResult->VEVENT->add(
                        'ATTENDEE',
                        'mailto:' . $deleteAttendee['email'],
                        array(
                            'CN' => !empty($deleteAttendee['name']) ? $deleteAttendee['name'] : $deleteAttendee['email'],
                            'PARTSTAT' => 'DECLINED'
                        )
                    );

                    Classes\Helper::sendAppointmentMessage($aData['UserPublicId'], $sEmail, $sSubject, $oVCalResult, 'REQUEST');
                }
            }
        }
        $oEvent->Attendees = @json_decode($aData['attendees'], true);
    }

    public function onProcessICSUpdateEvent(&$aData, &$mResult)
    {
        $oVCalResult = $aData['oVCalResult'];
        $oVEventResult = $aData['oVEventResult'];
        $sUserPublicId = $aData['sUserPublicId'];
        $sCalendarId = $aData['sCalendarId'];
        $sEventId = $aData['sEventId'];
        $sMethod = $aData['sMethod'];
        $sequence = $aData['sequence'];
        $sequenceServer = $aData['sequenceServer'];
        /** @var \Sabre\VObject\Component\VEvent $oVEvent */
        $oVEvent = $aData['oVEvent'];
        $mFromEmail = $aData['mFromEmail'];

        $sType = $sMethod;
        $oCurrentAttendee = null;
        if (isset($oVEvent->ATTENDEE) && $sequenceServer >= $sequence) {
            foreach ($oVEvent->ATTENDEE as $oAttendee) {
                if ($mFromEmail && $mFromEmail === str_replace('mailto:', '', strtolower((string) $oAttendee->getValue()))) {
                    $oCurrentAttendee = $oAttendee;
                    break;
                }
            }
            if (isset($oVEventResult->ATTENDEE, $oCurrentAttendee)) {
                foreach ($oVEventResult->ATTENDEE as &$oAttendeeResult) {
                    if ($oAttendeeResult->getValue() === $oCurrentAttendee->getValue()) {
                        if (isset($oCurrentAttendee['PARTSTAT'])) {
                            $oAttendeeResult['PARTSTAT'] = $oCurrentAttendee['PARTSTAT']->getValue();
                            $sType = $sType . '-' . (string) $oAttendeeResult['PARTSTAT'];
                            /** @var \Sabre\VObject\Property\ICalendar\DateTime $oLAST_MODIFIED */
                            $oLAST_MODIFIED = $oVEvent->{'LAST-MODIFIED'};

                            $oRespondedAt = $oLAST_MODIFIED->getDateTime();
                            $oRespondedAt->setTimezone(new \DateTimeZone('UTC'));
                            $oAttendeeResult['RESPONDED-AT'] = gmdate("Ymd\THis\Z", $oRespondedAt->getTimestamp());
                        }
                        break;
                    }
                }
            }
        }
        unset($oVCalResult->METHOD);
        $oVEventResult->{'LAST-MODIFIED'} = new \DateTime('now', new \DateTimeZone('UTC'));
        $mResult = $this->getManager()->updateEventRaw($sUserPublicId, $sCalendarId, $sEventId, $oVCalResult->serialize());
        if ($mResult) {
            $mResult = $sType;
        }
        //        $oVCalResult->METHOD = $sMethod;
    }

    public function onProcessICSCancel(&$aData, &$mResult)
    {
        $sCalendarId = $aData['sCalendarId'];
        $sUserPublicId = $aData['sUserPublicId'];
        $sEventId = $aData['sEventId'];

        if ($this->getManager()->deleteEvent($sUserPublicId, $sCalendarId, $sEventId)) {
            $mResult = true;
        }
    }

    public function onAddAttendeesToResult(&$aData, &$mResult)
    {
        $oVEventResult = $aData['oVEventResult'];
        $aAccountEmails = $aData['aAccountEmails'];
        $sMethod = $aData['sMethod'];

        if (isset($oVEventResult->ATTENDEE)) {
            foreach ($oVEventResult->ATTENDEE as $oAttendee) {
                $sAttendee = strtolower((string)$oAttendee);
                if (in_array($sAttendee, $aAccountEmails) && isset($oAttendee['PARTSTAT']) && $sMethod !== 'SAVE') {
                    $mResult['Attendee'] = str_replace('mailto:', '', $sAttendee);
                    $mResult['Action'] = $sMethod . '-' . $oAttendee['PARTSTAT']->getValue();
                }
            }
        }
    }

    public function onParseEvent(&$aData, &$aEvent)
    {
        $oVComponent = $aData['oVComponent'];
        $sOwnerEmail = $aData['sOwnerEmail'];
        $oUser = $aData['oUser'];

        $bIsAppointment = false;
        $aEvent['attendees'] = array();
        if (isset($oVComponent->ATTENDEE)) {
            $aEvent['attendees'] = \Aurora\Modules\Calendar\Classes\Parser::parseAttendees($oVComponent);

            if (isset($oVComponent->ORGANIZER)) {
                $sOwnerEmail = str_replace('mailto:', '', strtolower((string)$oVComponent->ORGANIZER));
                $iPos = strpos($sOwnerEmail, 'principals/');
                if ($iPos !== false) {
                    $sOwnerEmail = \trim(substr($sOwnerEmail, $iPos + 11), '/');
                }
                $aData['sOwnerEmail'] = $sOwnerEmail;
                $aEvent['organizer'] = $sOwnerEmail;
                $aEvent['organizerName'] = isset($oVComponent->ORGANIZER['CN']) ? (string)$oVComponent->ORGANIZER['CN'] : '';
            }
            $bIsAppointment = ($oUser instanceof \Aurora\Modules\Core\Models\User && $sOwnerEmail !== $oUser->PublicId);
        }

        $aEvent['appointment'] = $bIsAppointment;
        $aEvent['appointmentAccess'] = 0;
    }
}
