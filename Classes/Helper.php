<?php
/**
 * This code is licensed under AGPLv3 license or Afterlogic Software License
 * if commercial version of the product was purchased.
 * For full statements of the licenses see LICENSE-AFTERLOGIC and LICENSE-AGPL3 files.
 */

namespace Aurora\Modules\CalendarMeetingsPlugin\Classes;

use Aurora\Modules\Mail\Module as MailModule;
use Aurora\Modules\Core\Module as CoreModule;
use Aurora\System\Api;

/**
 * @license https://www.gnu.org/licenses/agpl-3.0.html AGPL-3.0
 * @license https://afterlogic.com/products/common-licensing Afterlogic Software License
 * @copyright Copyright (c) 2023, Afterlogic Corp.
 *
 * @internal
 *
 * @package Calendar
 * @subpackage Classes
 */
class Helper
{
    /**
     * @param string $sUserPublicId
     * @param string $sTo
     * @param string $sSubject
     * @param \Sabre\VObject\Component\VCalendar $oVCal
     * @param string $sMethod
     * @param string $sHtmlBody Default value is empty string.
     *
     * @throws \Aurora\System\Exceptions\ApiException
     *
     * @return \MailSo\Mime\Message
     */
    public static function sendAppointmentMessage($sUserPublicId, $sTo, $sSubject, $oVCal, $sMethod, $sHtmlBody = '', $oAccount = null, $sFromEmail = null)
    {
        // remove VALARM from invitation mail messages
        foreach ($oVCal->VEVENT as $key => $val) {
            if ($val->VALARM) {
                unset($oVCal->VEVENT[$key]->VALARM);
            }
        }
        $sBody = $oVCal->serialize();

        $oMessage = self::buildAppointmentMessage($sUserPublicId, $sTo, $sSubject, $sBody, $sMethod, $sHtmlBody, $oAccount, $sFromEmail);
        $oUser = CoreModule::Decorator()->GetUserByPublicId($sUserPublicId);
        if ($oUser instanceof \Aurora\Modules\Core\Models\User) {
            $oAccount = $oAccount ? $oAccount : MailModule::getInstance()->getAccountsManager()->getAccountUsedToAuthorize($oUser->PublicId);
            if ($oMessage && $oAccount instanceof \Aurora\Modules\Mail\Models\MailAccount) {
                try {
                    \Aurora\System\Api::Log('IcsAppointmentActionSendOriginalMailMessage');
                    if (MailModule::getInstance()->getMailManager()->sendMessage($oAccount, $oMessage)) {
                        $aCollection = $oMessage->GetRcpt();

                        $aEmails = [];
                        $aCollection->ForeachList(function ($oEmail) use (&$aEmails) {
                            $aEmails[strtolower($oEmail->GetEmail(true))] = trim($oEmail->GetDisplayName());
                        });

                        if (\is_array($aEmails)) {
                            $aArgs = [
                                'IdUser' => $oAccount->IdUser,
                                'Emails' => $aEmails
                            ];
                            $prevState = \Aurora\System\Api::skipCheckUserRole(true);
                            \Aurora\System\Api::GetModule('Mail')->broadcastEvent('AfterUseEmails', $aArgs);
                            \Aurora\System\Api::skipCheckUserRole($prevState);
                        }

                        return true;
                    }
                } catch (\Aurora\System\Exceptions\ManagerException $oException) {
                    throw new \Aurora\System\Exceptions\ApiException($oException->getCode(), $oException);
                }
            }
        }

        return false;
    }

    /**
     * @param string $sUserPublicId
     * @param string $sTo
     * @param string $sSubject
     * @param string $sBody
     * @param string $sMethod Default value is **null**.
     * @param string $sHtmlBody Default value is empty string.
     *
     * @return \MailSo\Mime\Message
     */
    public static function buildAppointmentMessage($sUserPublicId, $sTo, $sSubject, $sBody, $sMethod = null, $sHtmlBody = '', $oAccount = null, $sFromEmail = null)
    {
        $oMessage = null;
        $oUser = CoreModule::Decorator()->GetUserByPublicId($sUserPublicId);
        if (isset($sFromEmail)) {
            $sFrom = $sFromEmail;
        } else {
            $sFrom = $oAccount ? $oAccount->Email : $oUser->PublicId;
        }
        if ($oUser instanceof \Aurora\Modules\Core\Models\User && !empty($sTo) && !empty($sBody)) {
            $oMessage = \MailSo\Mime\Message::NewInstance();
            $oMessage->RegenerateMessageId();
            $oMessage->DoesNotCreateEmptyTextPart();

            $sXMailer = '';
            if (class_exists('\Aurora\Modules\Mail\Module')) {
                $oMailModule = \Aurora\Modules\Mail\Module::getInstance();
                $sXMailer = $oMailModule->oModuleSettings->XMailerValue;
            }

            if (0 < strlen($sXMailer)) {
                $oMessage->SetXMailer($sXMailer);
            }

            $oMessage
                ->SetFrom(\MailSo\Mime\Email::NewInstance($sFrom))
                ->SetSubject($sSubject)
            ;

            $oMessage->AddHtml($sHtmlBody);

            $oToEmails = \MailSo\Mime\EmailCollection::NewInstance($sTo);
            if ($oToEmails && $oToEmails->Count()) {
                $oMessage->SetTo($oToEmails);
            }

            if ($sMethod) {
                $oMessage->SetCustomHeader('Method', $sMethod);
            }

            $oMessage->AddAlternative(
                'text/calendar',
                \MailSo\Base\ResourceRegistry::CreateMemoryResourceFromString($sBody),
                \MailSo\Base\Enumerations\Encoding::_8_BIT,
                null === $sMethod ? array() : array('method' => $sMethod)
            );
        }

        return $oMessage;
    }

    /**
     * @param string $calendarId
     * @param string $eventId
     * @param string $sAccountEmail
     * @param string $sAttendee
     * @param string $sStartDate
     * @param string $location
     * @param string $description
     *
     * @return string
     */
    public static function createHtmlFromEvent($calendarId, $eventId, $sAccountEmail, $sAttendee, $sStartDate, $location, $description)
    {
        $sHtml = '';
        $aValues = [
            'attendee' => $sAttendee,
            'organizer' => $sAccountEmail,
            'calendarId' => $calendarId,
            'eventId' => $eventId
        ];

        $aValues['action'] = 'ACCEPTED';
        $sEncodedValueAccept = \Aurora\System\Api::EncodeKeyValues($aValues);
        $aValues['action'] = 'TENTATIVE';
        $sEncodedValueTentative = \Aurora\System\Api::EncodeKeyValues($aValues);
        $aValues['action'] = 'DECLINED';
        $sEncodedValueDecline = \Aurora\System\Api::EncodeKeyValues($aValues);

        $fullUrl = \MailSo\Base\Http::SingletonInstance()->GetFullUrl();

        /** @var \Aurora\Modules\Dav\Module */
        $oDavModule = Api::GetModule('Dav');
        if ($oDavModule && $oDavModule->getModuleSettings()->ProductUrlForExternalClients !== '') {
            $fullUrl = $oDavModule->getModuleSettings()->ProductUrlForExternalClients;
        }

        $sHref = rtrim($fullUrl, '\\/ ') . '/?invite=';
        $oCalendarMeetingsModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
        if ($oCalendarMeetingsModule instanceof \Aurora\System\Module\AbstractModule) {
            $sHtml = file_get_contents($oCalendarMeetingsModule->GetPath() . '/templates/CalendarEventInvite.html');
            $sHtml = strtr($sHtml, [
                '{{INVITE/LOCATION}}'	=> $oCalendarMeetingsModule->i18N('LOCATION'),
                '{{INVITE/WHEN}}'		=> $oCalendarMeetingsModule->I18N('WHEN'),
                '{{INVITE/DESCRIPTION}}' => $oCalendarMeetingsModule->i18N('DESCRIPTION'),
                '{{INVITE/INFORMATION}}' => $oCalendarMeetingsModule->i18N('INFORMATION', ['Email' => $sAttendee]),
                '{{INVITE/ACCEPT}}'		=> $oCalendarMeetingsModule->i18N('ACCEPT'),
                '{{INVITE/TENTATIVE}}'	=> $oCalendarMeetingsModule->i18N('TENTATIVE'),
                '{{INVITE/DECLINE}}'	=> $oCalendarMeetingsModule->i18N('DECLINE'),
                '{{Location}}'			=> $location,
                '{{Start}}'				=> $sStartDate,
                '{{Description}}'		=> $description,
                '{{HrefAccept}}'		=> $sHref . $sEncodedValueAccept,
                '{{HrefTentative}}'		=> $sHref . $sEncodedValueTentative,
                '{{HrefDecline}}'		=> $sHref . $sEncodedValueDecline
            ]);
        }

        return $sHtml;
    }

    /**
     * @param string $sUserPublicId
     * @param string $sTo
     * @param string $sSubject
     * @param string $sHtmlBody
     *
     * @throws \Aurora\System\Exceptions\ApiException
     *
     * @return \MailSo\Mime\Message
     */
    public static function sendSelfNotificationMessage($sUserPublicId, $sTo, $sSubject, $sHtmlBody)
    {
        $oMessage = self::buildSelfNotificationMessage($sUserPublicId, $sTo, $sSubject, $sHtmlBody);
        $oUser = CoreModule::Decorator()->GetUserByPublicId($sUserPublicId);
        if ($oUser instanceof \Aurora\Modules\Core\Models\User) {
            $oAccount = MailModule::getInstance()->getAccountsManager()->getAccountUsedToAuthorize($oUser->PublicId);
            if ($oMessage && $oAccount instanceof \Aurora\Modules\Mail\Models\MailAccount) {
                try {
                    \Aurora\System\Api::Log('IcsAppointmentActionSendSelfMailMessage');
                    return MailModule::getInstance()->getMailManager()->sendMessage($oAccount, $oMessage);
                } catch (\Aurora\System\Exceptions\ManagerException $oException) {
                    $iCode = \Aurora\Modules\Mail\Enums\ErrorCodes::CannotSendMessage;
                    switch ($oException->getCode()) {
                        case \Aurora\Modules\Mail\Enums\ErrorCodes::CannotSendMessageInvalidRecipients:
                            $iCode = \Aurora\Modules\Mail\Enums\ErrorCodes::CannotSendMessageInvalidRecipients;
                            break;
                    }

                    throw new \Aurora\System\Exceptions\ApiException($iCode, $oException);
                }
            }
        }

        return false;
    }

    /**
     * @param string $sUserPublicId
     * @param string $sTo
     * @param string $sSubject
     * @param string $sHtmlBody
     *
     * @return \MailSo\Mime\Message
     */
    public static function buildSelfNotificationMessage($sUserPublicId, $sTo, $sSubject, $sHtmlBody)
    {
        $oMessage = null;
        $oUser = CoreModule::Decorator()->GetUserByPublicId($sUserPublicId);
        if ($oUser instanceof \Aurora\Modules\Core\Models\User && !empty($sTo) && !empty($sHtmlBody)) {
            $oMessage = \MailSo\Mime\Message::NewInstance();
            $oMessage->RegenerateMessageId();
            $oMessage->DoesNotCreateEmptyTextPart();

            $sXMailer = '';
            if (class_exists('\Aurora\Modules\Mail\Module')) {
                $oMailModule = \Aurora\Modules\Mail\Module::getInstance();
                $sXMailer = $oMailModule->oModuleSettings->XMailerValue;
            }

            if (0 < strlen($sXMailer)) {
                $oMessage->SetXMailer($sXMailer);
            }

            $oMessage
                ->SetFrom(\MailSo\Mime\Email::NewInstance($oUser->PublicId))
                ->SetSubject($sSubject)
            ;

            $oMessage->AddHtml($sHtmlBody);

            $oToEmails = \MailSo\Mime\EmailCollection::NewInstance($sTo);
            if ($oToEmails && $oToEmails->Count()) {
                $oMessage->SetTo($oToEmails);
            }
        }

        return $oMessage;
    }

    public static function createSelfNotificationSubject($sAction, $sEventName)
    {
        $sResult = self::getSelfNotificationActionName($sAction) . ': ' . $sEventName;

        return $sResult;
    }

    public static function createSelfNotificationHtmlBody($sAction, $aEvent, $sEmail, $sCalendarName, $sStartDate)
    {
        $sHtml = '';
        $sActionName = self::getSelfNotificationActionName($sAction);
        $oCalendarMeetingsModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
        if ($oCalendarMeetingsModule instanceof \Aurora\System\Module\AbstractModule) {
            $sHtml = file_get_contents($oCalendarMeetingsModule->GetPath() . '/templates/CalendarEventSelfNotification.html');
            $sHtml = strtr($sHtml, [
                '{{LOCATION}}'			=> $oCalendarMeetingsModule->i18N('LOCATION'),
                '{{WHEN}}'				=> $oCalendarMeetingsModule->I18N('WHEN'),
                '{{DESCRIPTION}}'		=> $oCalendarMeetingsModule->i18N('DESCRIPTION'),
                '{{INFORMATION}}'		=> $oCalendarMeetingsModule->i18N('INFORMATION', ['Email' => $sEmail]),
                '{{REACTION}}'			=> $oCalendarMeetingsModule->i18N('USER_REACTION'),
                '{{Calendar}}'			=> $sCalendarName . ' ' . $sEmail,
                '{{Location}}'			=> $aEvent['location'],
                '{{Start}}'				=> $sStartDate,
                '{{Description}}'			=> $aEvent['description'],
                '{{Reaction}}'			=> $sActionName
            ]);
        }

        return $sHtml;
    }

    public static function getSelfNotificationActionName($sAction)
    {
        $sResult = '';
        $oCalendarMeetingsModule = \Aurora\System\Api::GetModule('CalendarMeetingsPlugin');
        switch ($sAction) {
            case 'ACCEPTED':
                $sResult = $oCalendarMeetingsModule->i18N('ACCEPT');
                break;
            case 'DECLINED':
                $sResult = $oCalendarMeetingsModule->i18N('DECLINE');
                break;
            case 'TENTATIVE':
                $sResult = $oCalendarMeetingsModule->i18N('TENTATIVE');
                break;
        }

        return $sResult;
    }
}
