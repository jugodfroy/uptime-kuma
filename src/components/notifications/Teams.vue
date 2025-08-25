<template>
    <div class="mb-3">
        <label for="teams-webhookurl" class="form-label">{{ $t("Webhook URL") }}</label>
        <input
            id="teams-webhookurl"
            v-model="$parent.notification.webhookUrl"
            type="text"
            class="form-control"
            required
        />
        <i18n-t tag="div" keypath="wayToGetTeamsURL" class="form-text">
            <a
                href="https://support.microsoft.com/en-us/office/create-incoming-webhooks-with-workflows-for-microsoft-teams-8ae491c7-0394-4861-ba59-055e33f75498"
                target="_blank"
            >{{ $t("here") }}</a>
        </i18n-t>
    </div>
</template>
