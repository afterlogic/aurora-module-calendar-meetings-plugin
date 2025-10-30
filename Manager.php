<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin;

use Aurora\Modules\Core\Models\User;
use Aurora\Modules\Mail\Models\MailAccount;
use Aurora\System\Api;
use Aurora\Modules\Core\Module as CoreModule;
use Aurora\System\Enums\LogLevel;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2023, Afterlogic Corp.
 */
class Manager extends \Aurora\Modules\Calendar\Manager
{
    /**
     * Processing response to event invitation. [Aurora only.](http://dev.afterlogic.com/aurora)
     *
     * @param string $sUserPublicId
     * @param string $sCalendarId Calendar ID
     * @param string $sEventId Event ID
     * @param string $sAttendee Attendee identified by email address
     * @param string $sAction Appointment actions. Accepted values:
     *		- "ACCEPTED"
     *		- "DECLINED"
     *		- "TENTATIVE"
     *
     * @return bool
     */
    public function updateAppointment($sUserPublicId, $sCalendarId, $sEventId, $sAttendee, $sAction)
    {
        $oResult = null;

        $aData = $this->oStorage->getEvent($sUserPublicId, $sCalendarId, $sEventId);
        if ($aData !== false) {
            $oVCal = $aData['vcal'];
            $oVCal->METHOD = 'REQUEST';
            return $this->appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $oVCal);
        }

        return $oResult;
    }

    /**
     * @param string $sPartstat
     * @param string $sSummary
     */
    protected function getMessageSubjectFromPartstat($sPartstat, $sSummary)
    {
        $sSubject = $sSummary;
        switch ($sPartstat) {
            case 'ACCEPTED':
                $sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_ACCEPTED') . ': ' . $sSummary;
                break;
            case 'DECLINED':
                $sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_DECLINED') . ': ' . $sSummary;
                break;
            case 'TENTATIVE':
                $sSubject = $this->GetModule()->i18N('SUBJECT_PREFFIX_TENTATIVE') . ': ' . $sSummary;
                break;
        }

        return $sSubject;
    }

    /**
     * @param User $oUser
     * @param string $sAttendee
     */
    protected function getFromAccount($oUser, $sAttendee)
    {
        $oFromAccount = null;

        if ($oUser && $oUser->PublicId !== $sAttendee) {
            $oMailModule = Api::GetModule('Mail');
            /** @var \Aurora\Modules\Mail\Module $oMailModule */
            if ($oMailModule) {
                $aAccounts = $oMailModule->getAccountsManager()->getUserAccounts($oUser->Id);
                foreach ($aAccounts as $oAccount) {
                    if ($oAccount instanceof MailAccount && $oAccount->Email === $sAttendee) {
                        $oFromAccount = $oAccount;
                        break;
                    }
                }
            }
        }

        return $oFromAccount;
    }

    /**
     * Allows for responding to event invitation (accept / decline / tentative). [Aurora only.](http://dev.afterlogic.com/aurora)
     *
     * @param int|string $sUserPublicId Account object
     * @param string $sAttendee Attendee identified by email address
     * @param string $sAction Appointment actions. Accepted values:
     *		- "ACCEPTED"
     *		- "DECLINED"
     *		- "TENTATIVE"
     * @param string $sCalendarId Calendar ID
     * @param \Sabre\VObject\Component\VCalendar $oVCal VCalendar object with ICS data of the response
     * @param int $AllEvents  Indicator whether the response is for all events or a single event:
     *		- 0 - single event
     *		- 1 - all events in the series
     *		- 2 - all events including future ones
     * @param int|null $RecurrenceId Recurrence ID for single event response
     * @param bool $bIsLinkAction If **true**, the attendee's action on the link is assumed
     * @param bool $bIsExternalAttendee
     *
     * @return bool
     */
    public function appointmentAction($sUserPublicId, $sAttendee, $sAction, $sCalendarId, $oVCal, $AllEvents = 2, $RecurrenceId = null, $bIsLinkAction = false, $bIsExternalAttendee = false)
    {
        $oUser = null;
        $bResult = false;
        $sEventId = null;
        $sTo = $sSubject = '';

        $oUser = CoreModule::Decorator()->GetUserByPublicId($sUserPublicId);

        if (!($oUser instanceof User)) {
            throw new Exceptions\Exception(
                Enums\ErrorCodes::CannotSendAppointmentMessage,
                null,
                'User not found'
            );
        }

        if ($bIsLinkAction && !$bIsExternalAttendee) {
            // getting default calendar for attendee
            $oCalendar = $this->getDefaultCalendar($oUser->PublicId);
            if ($oCalendar) {
                $sCalendarId = $oCalendar->Id;
            }
        }

        if ($oVCal) {
            $sMethod = strtoupper((string) $oVCal->METHOD);
            $sPartstat = strtoupper($sAction);
            $sCN = '';
            if ($sAttendee ===  $oUser->PublicId) {
                $sCN = !empty($oUser->Name) ? $oUser->Name : $sAttendee;
            }
            $bNeedsToUpdateEvent = false;

            // Now we need to loop through the original organizer event, to find
            // all the instances where we have a reply for.
            $event = $oVCal->getBaseComponent('VEVENT');

            if (!$event) {
                // No master event
                $event = $oVCal->VEVENT;
                $oRecurrenceId = $event->{'RECURRENCE-ID'};
                if ($oRecurrenceId) {
                    $AllEvents = 1;
                    $dRecurrence = $oRecurrenceId->getDateTime();
                    $RecurrenceId = $dRecurrence->getTimestamp();
                } else {
                    return false;
                }
            }

            $sEventId = (string) $event->UID;
            if ($AllEvents === 2) {
                if ($sPartstat === 'DECLINED' || $sMethod === 'CANCEL') {
                    if ($sCalendarId !== false) {
                        $this->deleteEvent($sAttendee, $sCalendarId, $sEventId);
                    }
                } else {
                    $bNeedsToUpdateEvent = true;
                }
            }

            $foundedAttendee = null;
            if (isset($event->ATTENDEE)) {
                foreach ($event->ATTENDEE as $attendee) {
                    $sEmail = str_replace('mailto:', '', strtolower($attendee->getValue()));
                    if (strtolower($sEmail) === strtolower($sAttendee)) {
                        $attendee['PARTSTAT'] = $sPartstat;
                        $attendee['RESPONDED-AT'] = gmdate("Ymd\THis\Z");
                        // Un-setting the RSVP status, because we now know
                        // that the attendee already replied.
                        unset($attendee['RSVP']);
                        $foundedAttendee = $attendee;
                        break;
                    }
                }
            }

            unset($event->ATTENDEE);

            if($foundedAttendee) {
                $event->add($foundedAttendee);
            } else {
                // Adding a new attendee. The iTip documentation calls this
                // a party crasher.
                $attendee = $event->add('ATTENDEE', 'mailto:' . $sAttendee, [
                    'PARTSTAT' => $sPartstat,
                    'CN' => $sCN,
                    'RESPONDED-AT' => gmdate("Ymd\THis\Z")
                ]);
            }

            $event->{'LAST-MODIFIED'} = new \DateTime('now', new \DateTimeZone('UTC'));

            if ($AllEvents === 1 && $RecurrenceId !== null) {
                if ($sCalendarId !== false) {
                    $oEvent = new \Aurora\Modules\Calendar\Classes\Event();
                    $oEvent->IdCalendar = $sCalendarId;
                    $oEvent->Id = $sEventId;
                    $bDelete = ($sPartstat === 'DECLINED' || $sMethod === 'CANCEL');

                    $oEvent->populateFromVEvent($sUserPublicId, $event);

                    $this->updateExclusion($sAttendee, $oEvent, $RecurrenceId, $bDelete);
                }

                $vevent = null;
                $index = \Aurora\Modules\Calendar\Classes\Helper::isRecurrenceExists($oVCal->VEVENT, $RecurrenceId);
                if ($index === false) {
                    // If we got replies to instances that did not exist in the
                    // original list, it means that new exceptions must be created.
                    $recurrenceIterator = new \Sabre\VObject\Recur\EventIterator($oVCal, $event->UID);
                    $found = false;
                    $iterations = 1000;
                    do {
                        $newObject = $recurrenceIterator->getEventObject();
                        $recurrenceIterator->next();

                        if (isset($newObject->{'RECURRENCE-ID'})) {
                            $iRecurrenceId = \Aurora\Modules\Calendar\Classes\Helper::getTimestamp($newObject->{'RECURRENCE-ID'}, $oUser->DefaultTimeZone);
                            if ((int) $iRecurrenceId === (int) $RecurrenceId) {
                                $found = true;
                            }
                        }
                        --$iterations;
                    } while ($recurrenceIterator->valid() && !$found && $iterations);

                    if ($found) {
                        unset(
                            $newObject->RRULE,
                            $newObject->EXDATE,
                            $newObject->RDATE
                        );
                        $vevent = $oVCal->add($newObject);
                    }
                } else {
                    $vevent = $oVCal->VEVENT[$index];
                }

                $foundedAttendee = null;
                if (isset($vevent->ATTENDEE)) {
                    foreach ($vevent->ATTENDEE as $attendee) {
                        $sEmail = str_replace('mailto:', '', strtolower($attendee->getValue()));
                        if (strtolower($sEmail) === strtolower($sAttendee)) {
                            $attendee['PARTSTAT'] = $sPartstat;
                            $attendee['RESPONDED-AT'] = gmdate("Ymd\THis\Z");
                            $foundedAttendee = $attendee;
                            break;
                        }
                    }
                }
                if ($vevent) {
                    unset($vevent->ATTENDEE);
                    if ($foundedAttendee) {
                        $vevent->add($foundedAttendee);
                    } else {
                        // Adding a new attendee
                        $attendee = $vevent->add('ATTENDEE', 'mailto:' . $sAttendee, [
                            'PARTSTAT' => $sPartstat,
                            'CN' => $sCN,
                            'RESPONDED-AT' => gmdate("Ymd\THis\Z")
                        ]);
                    }
                }
            }

            if ($sMethod === 'REQUEST') {
                $sMethod = 'REPLY';
            }

            $oVCalForSend = clone $oVCal;
            $oVCalForSend->METHOD = $sMethod;

            $sTo = isset($event->ORGANIZER) ? str_replace(['mailto:', 'principals/'], '', strtolower((string) $event->ORGANIZER)) : '';
            $sSummary = isset($event->SUMMARY) ? (string) $event->SUMMARY : '';
            $sSubject = $this->getMessageSubjectFromPartstat($sPartstat, $sSummary);

            if ($bNeedsToUpdateEvent) { // update event on server
                unset($oVCal->METHOD);
                $this->oStorage->updateEventRaw(
                    $oUser->PublicId,
                    $sCalendarId,
                    $sEventId,
                    $oVCal->serialize()
                );
            }

            if ($oVCalForSend) { //send message to organizer
                if (empty($sTo)) {
                    throw new Exceptions\Exception(
                        Enums\ErrorCodes::CannotSendAppointmentMessageNoOrganizer
                    );
                }
                $oFromAccount = $this->getFromAccount($oUser, $sAttendee);
                $bResult = Classes\Helper::sendAppointmentMessage(
                    $oUser->PublicId,
                    $sTo,
                    $sSubject,
                    $oVCalForSend,
                    $sMethod,
                    '',
                    $oFromAccount,
                    $sAttendee
                );
            }
        }

        if (!$bResult) {
            Api::Log('Ics Appointment Action FALSE result!', LogLevel::Error);
            if ($sUserPublicId) {
                Api::Log('Email: ' . $oUser->PublicId . ', Action: ' . $sAction . ', Data:', LogLevel::Error);
            }
            if ($oVCal instanceof \Sabre\VObject\Component\VCalendar) {
                Api::Log($oVCal->serialize(), LogLevel::Error);
            }
        } else {
            $bResult = $sEventId;
        }

        return $bResult;
    }
}
