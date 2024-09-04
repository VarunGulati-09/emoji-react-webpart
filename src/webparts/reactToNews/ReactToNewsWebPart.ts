import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI } from "@pnp/sp";

import styles from './ReactToNewsWebPart.module.scss';

export interface IReactToNewsWebPartProps {
  description: string;
}

interface IReactionItem {
  Id: number;
  Reaction: string;
  User: string;
}

export default class ReactToNewsWebPart extends BaseClientSideWebPart<IReactToNewsWebPartProps> {
  private reactionCounts: { [key: string]: number } = {};
  private userReaction: string | null = null;
  private userReactionId: number | null = null;
  private sp: SPFI;
  private fallbackPostId: string = 'default-post-id';
  private currentUser: string = '';

  public onInit(): Promise<void> {
    return super.onInit().then(async _ => {
      this.sp = spfi().using(SPFx(this.context as WebPartContext));
      const user = await this.sp.web.currentUser();
      this.currentUser = user.Title;
      return this.loadReactions();
    });
  }

  public render(): void {
    if (!this.domElement.innerHTML) {
      this.domElement.innerHTML = `
        <div class="${styles.reactToNews}">
          <div class="${styles.container}">
            <div class="${styles.reactionContainer}">
              <span class="${styles.defaultReaction}" data-reaction="like">👍</span>
              <div class="${styles.reactionOptions}">
                <span class="${styles.emoji}" data-reaction="like">👍</span>
                <span class="${styles.emoji}" data-reaction="love">❤️</span>
                <span class="${styles.emoji}" data-reaction="laugh">😂</span>
                <span class="${styles.emoji}" data-reaction="surprised">😮</span>
                <span class="${styles.emoji}" data-reaction="sad">😢</span>
                <span class="${styles.emoji}" data-reaction="angry">😡</span>
              </div>
            </div>
            <div class="${styles.reactionCount}"></div>
          </div>
        </div>`;
      this.attachEventListeners();
    }
    this.updateReactionCount();
  }

  private updateReactionCount(): void {
    const countElement = this.domElement.querySelector(`.${styles.reactionCount}`);
    if (countElement) {
      countElement.innerHTML = this.renderReactionCount();
    }
  }

  private renderReactionCount(): string {
    const totalReactions = Object.keys(this.reactionCounts).reduce((sum, key) => sum + this.reactionCounts[key], 0);
    if (totalReactions === 0) return '';

    const reactionStrings = Object.keys(this.reactionCounts)
      .map(reaction => `${this.getEmojiForReaction(reaction)} ${this.reactionCounts[reaction]}`)
      .join(' ');

    return `${reactionStrings}`;
  }

  private getEmojiForReaction(reaction: string): string {
    const emojiMap: { [key: string]: string } = {
      'like': '👍', 'love': '❤️', 'laugh': '😂',
      'surprised': '😮', 'sad': '😢', 'angry': '😡'
    };
    return emojiMap[reaction] || '👍';
  }

  private attachEventListeners(): void {
    const emojis = this.domElement.querySelectorAll(`.${styles.emoji}`);
    emojis.forEach((emoji: HTMLElement) => {
      emoji.addEventListener('click', this.handleReactionClick.bind(this));
    });
  }

  private async handleReactionClick(event: MouseEvent): Promise<void> {
    const target = event.currentTarget as HTMLElement;
    const reaction = target.getAttribute('data-reaction');

    if (!reaction) return;

    if (this.userReaction === reaction) {
      // Remove reaction if clicking the same one
      await this.removeUserReaction();
    } else {
      // Change reaction
      await this.saveUserReaction(reaction);
    }

    await this.loadReactions();
    this.render();
  }

  private async saveUserReaction(reaction: string): Promise<void> {
    const postId = this.context.pageContext.listItem?.id?.toString() || this.fallbackPostId;

    try {
      if (this.userReactionId) {
        // Update existing reaction
        await this.sp.web.lists.getByTitle('Reactions').items.getById(this.userReactionId).update({
          Reaction: reaction
        });
      } else {
        // Add new reaction
        await this.sp.web.lists.getByTitle('Reactions').items.add({
          Title: postId,
          User: this.currentUser,
          Reaction: reaction
        });
      }
      console.log('Reaction saved successfully');
    } catch (error) {
      console.error('Error saving reaction:', error);
    }
  }

  private async removeUserReaction(): Promise<void> {
    if (this.userReactionId) {
      try {
        await this.sp.web.lists.getByTitle('Reactions').items.getById(this.userReactionId).delete();
        console.log('Reaction removed successfully');
        this.userReaction = null;
        this.userReactionId = null;
      } catch (error) {
        console.error('Error removing reaction:', error);
      }
    }
  }

  private async loadReactions(): Promise<void> {
    const postId = this.context.pageContext.listItem?.id?.toString() || this.fallbackPostId;

    try {
      const items: IReactionItem[] = await this.sp.web.lists.getByTitle('Reactions').items.filter(`Title eq '${postId}'`)();

      // Reset reaction counts
      this.reactionCounts = {};
      this.userReaction = null;
      this.userReactionId = null;

      items.forEach(item => {
        const reaction = item.Reaction;
        this.reactionCounts[reaction] = (this.reactionCounts[reaction] || 0) + 1;

        if (item.User === this.currentUser) {
          this.userReaction = reaction;
          this.userReactionId = item.Id;
        }
      });

      console.log('Loaded reactions:', this.reactionCounts);
    } catch (error) {
      console.error('Error loading reactions:', error);
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "React to News Web Part"
          },
          groups: [
            {
              groupName: "Settings",
              groupFields: [
                PropertyPaneTextField('description', {
                  label: "Description"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}